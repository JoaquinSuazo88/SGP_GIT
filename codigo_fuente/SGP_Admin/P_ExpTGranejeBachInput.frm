VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form P_ExpTGranejeBachInput 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportar Excel tabla gramaje & Bach - Input"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   10155
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
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
      Height          =   495
      Index           =   1
      Left            =   8040
      TabIndex        =   8
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   5880
      TabIndex        =   7
      Top             =   6000
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   9975
      Begin VB.Frame Frame2 
         Height          =   1455
         Left            =   6120
         TabIndex        =   22
         Top             =   1560
         Width           =   3255
         Begin VB.OptionButton Option2 
            Caption         =   "Sólo las líneas con cambio de las recetas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   240
            TabIndex        =   24
            Top             =   840
            Width           =   2295
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Todas las recetas y sus líneas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   240
            TabIndex        =   23
            Top             =   240
            Value           =   -1  'True
            Width           =   2295
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Filtro Ingrediente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3480
         TabIndex        =   21
         Top             =   1680
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Filtro C.Dietica y Tipo Plato"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   0
         Left            =   960
         TabIndex        =   20
         Top             =   1680
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Bach - Input"
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
         Left            =   360
         TabIndex        =   14
         Top             =   3360
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
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
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   9
         Top             =   480
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   4320
         Width           =   7215
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Left            =   1755
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   3840
         Width           =   7335
         _Version        =   196608
         _ExtentX        =   12938
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
         ButtonStyle     =   2
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
         NoSpecialKeys   =   1
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   0
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
         ControlType     =   3
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
      Begin MSComctlLib.ProgressBar prbStatus 
         Height          =   285
         Left            =   360
         TabIndex        =   3
         Top             =   5280
         Width           =   8670
         _ExtentX        =   15293
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComDlg.CommonDialog CD 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   1
         Left            =   2505
         TabIndex        =   10
         Top             =   825
         Width           =   915
         _Version        =   196608
         _ExtentX        =   1614
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
         AutoCase        =   0
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
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   1
         Left            =   2520
         TabIndex        =   15
         Top             =   1185
         Width           =   915
         _Version        =   196608
         _ExtentX        =   1614
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
         AlignTextV      =   0
         AllowNull       =   -1  'True
         NoSpecialKeys   =   3
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
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   ""
         MaxValue        =   "2147483647"
         MinValue        =   "-2147483647"
         NegFormat       =   0
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483637
         Appearance      =   1
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483637
         AutoMenu        =   -1  'True
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Opcional"
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
         Left            =   9120
         TabIndex        =   19
         Top             =   1280
         Width           =   765
      End
      Begin VB.Line Line1 
         X1              =   360
         X2              =   9615
         Y1              =   3120
         Y2              =   3135
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   3840
         TabIndex        =   17
         Top             =   1185
         Width           =   5175
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   4
         Left            =   3390
         Picture         =   "P_ExpTGranejeBachInput.frx":0000
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Regimen"
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
         Left            =   960
         TabIndex        =   16
         Top             =   1200
         Width           =   750
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   9
         Left            =   3825
         TabIndex        =   12
         Top             =   840
         Width           =   5175
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   6
         Left            =   3375
         Picture         =   "P_ExpTGranejeBachInput.frx":030A
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label2 
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
         Index           =   4
         Left            =   960
         TabIndex        =   11
         Top             =   885
         Width           =   1380
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Archivo Origen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   6
         Top             =   3915
         Width           =   1275
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "lblStatus"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   5040
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hoja"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   4440
         Width           =   405
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   10
         Left            =   3840
         TabIndex        =   13
         Top             =   870
         Width           =   5205
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   5
         Left            =   3855
         TabIndex        =   18
         Top             =   1200
         Width           =   5205
      End
   End
End
Attribute VB_Name = "P_ExpTGranejeBachInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Public lc_Aux As String

Private obj_Excel       As Object
Private obj_Workbook    As Object
Private obj_Worksheet   As Object

Dim FilCatDie As Long
Dim FilTipPla As Long
Dim MsgTitulo As String

Private Sub Command1_Click(Index As Integer)

On Error GoTo Man_Error

Select Case Index
    
    Case 0
    
        If ValidaDatos = False Then Exit Sub
           
        If Option1(0).Value = True Then
        
            ExportarPlanillaExcel
        
        ElseIf Option1(1).Value = True Then
        
            If MsgBox("Esta seguro realizar cambio...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
            
            ActualizarCambioTablaGramaje
        
        End If
    
    Case 1
    
        Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Salir"), Me.HelpContextID, "", "", "")
        Unload Me
    
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub ExportarPlanillaExcel()

On Error GoTo Man_Error

Dim RS              As New ADODB.Recordset
Dim Sql             As String
Dim NomArchivoExcel As String
Dim Extension       As String

Dim xlApp           As Object
Dim xlWb            As Object
Dim xlWs            As Object
  
Command1(0).Enabled = False
  
'-------> Guardar nombre archivo excel
NomArchivoExcel = ""
CD.Flags = &H80000 Or &H400& Or &H1000& Or &H4&
CD.Filter = "Todos los archivos *.xls,*.xlsx"
On Error Resume Next
CD.ShowSave
           
 '-------> JPAZ Permite controlar Boton Cancelar
 If Err.Number = 32755 Then
    
    MsgBox "Proceso cancelado"
    Exit Sub
 
 End If
            
 If CD.FileName = "" Then
    
    MsgBox "Debe seleccionar la ruta y nombre de archivo", vbExclamation
    Exit Sub
 
 Else
    
    Extension = ""
    Extension = Right(CD.FileName, Len(CD.FileName) - (InStrRev(CD.FileName, ".")))
    If UCase(Extension) <> "XLS" And UCase(Extension) <> "XLSX" Then
       
       MsgBox "La extensión del archivo debe ser (*.xls,*.xlsx)", vbCritical
       Exit Sub
    
    End If
    NomArchivoExcel = CD.FileName
 
 End If
          
fg_carga ""

Dim XmlDietetica   As String
Dim XmlPlato       As String
Dim XmlIngrediente As String
Dim IndFiltro      As Long
Dim i              As Long
Dim Die            As Long
Dim Pla            As Long
Dim CodIng         As String

'---------> Armar Xml Categoria Dietetica , Tipo Plato & Ingredientes

Let XmlDietetica = ""
Let XmlDietetica = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
Let XmlDietetica = XmlDietetica & "<Dietetica>"

For IndFiltro = 1 To B_DieTipExcel.TvwDir(0).Nodes.count

    If B_DieTipExcel.TvwDir(0).Nodes.item(IndFiltro).Checked = True And Trim(B_DieTipExcel.TvwDir(0).Nodes.item(IndFiltro).text) <> "*" Then
          
       XmlDietetica = XmlDietetica & " <DetDietetica"
       
       XmlDietetica = XmlDietetica & " Die = " & Chr(34) & Val(Mid(B_DieTipExcel.TvwDir(0).Nodes.item(IndFiltro).text, 1, InStr(B_DieTipExcel.TvwDir(0).Nodes.item(IndFiltro).text, " - ") - 1)) & Chr(34)
       XmlDietetica = XmlDietetica & "/>"

    End If
       
Next IndFiltro

XmlDietetica = XmlDietetica & "</Dietetica>"

Let XmlPlato = ""
Let XmlPlato = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
Let XmlPlato = XmlPlato & "<Plato>"

For IndFiltro = 1 To B_DieTipExcel.TvwDir(1).Nodes.count

    If B_DieTipExcel.TvwDir(1).Nodes.item(IndFiltro).Checked = True And Trim(B_DieTipExcel.TvwDir(1).Nodes.item(IndFiltro).text) <> "*" Then
          
       XmlPlato = XmlPlato & " <DetPlato"
       
       XmlPlato = XmlPlato & " Pla = " & Chr(34) & Val(Mid(B_DieTipExcel.TvwDir(1).Nodes.item(IndFiltro).text, 1, InStr(B_DieTipExcel.TvwDir(1).Nodes.item(IndFiltro).text, "-") - 1)) & Chr(34)
       XmlPlato = XmlPlato & "/>"

    End If
       
Next IndFiltro

XmlPlato = XmlPlato & "</Plato>"

Let XmlIngrediente = ""
Let XmlIngrediente = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
Let XmlIngrediente = XmlIngrediente & "<Ingred>"

For i = 1 To B_TablaEstandar.vaSpread1.MaxRows

    B_TablaEstandar.vaSpread1.Row = i
    B_TablaEstandar.vaSpread1.Col = 1
    If B_TablaEstandar.vaSpread1.RowHidden = False And B_TablaEstandar.vaSpread1.text = "1" Then
          
       XmlIngrediente = XmlIngrediente & " <DetIng"
       
       B_TablaEstandar.vaSpread1.Col = 2
       CodIng = B_TablaEstandar.vaSpread1.text
       
       XmlIngrediente = XmlIngrediente & " Ing = " & Chr(34) & CodIng & Chr(34)
       XmlIngrediente = XmlIngrediente & "/>"

    End If
       
Next i

XmlIngrediente = XmlIngrediente & "</Ingred>"

'-------> Validar cantidad registro se sobre pase hoja excel
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute("sgpadm_Sel_ExportartablaGramajeCeco_V02 '" & fpText(1).text & "', " & IIf(Trim(fpLongInteger1(1).text) <> "", fpLongInteger1(1).text, 0) & ", '" & XmlDietetica & "', '" & XmlPlato & "', '" & XmlIngrediente & "', '" & IIf(Option2(0).Value = True, "1", "2") & "'")
If Not RS.EOF Then
  
   If RS.RecordCount > 1020000 And UCase(Extension) = "XLSX" Then
      
      Command1(0).Enabled = True
        
      '-------> Close ADO objects
      RS.Close
      Set RS = Nothing
      fg_descarga
      MsgBox "El resultado sobrepasa maximo de fila en excel 1020000, proceso cancelado utilice filtro categoria dietetica o bien tipo de plato", vbCritical
      Exit Sub
   
   ElseIf UCase(Extension) = "XLS" And RS.RecordCount > 65533 Then
   
      Command1(0).Enabled = True
      
      '-------> Close ADO objects
      RS.Close
      Set RS = Nothing
      
      MsgBox "El resultado sobrepasa maximo de fila en excel 65533, proceso cancelado utilice filtro categoria dietetica o bien tipo de plato", vbCritical
      Exit Sub
   
   
   End If
  
End If

'-------> Create an instance of Excel and add a workbook
Set xlApp = CreateObject("Excel.Application")
Set xlWb = xlApp.Workbooks.Add
Set xlWs = xlWb.Worksheets("Hoja1")
  
'-------> Display Excel and give user control of Excel's lifetime
'    xlApp.Visible = True
xlApp.UserControl = True

'------> desactiva los mensaje
xlApp.DisplayAlerts = False

'-------> Check version of Excel
Call encabezado(RS, xlWs)
          
xlWs.Cells(2, 1).CopyFromRecordset RS
'-------> Auto-fit the column widths and row heights
xlApp.Selection.CurrentRegion.Columns.AutoFit
xlApp.Selection.CurrentRegion.Rows.AutoFit
    
'xlApp.Columns("A:A").Select
'xlApp.Selection.Delete Shift:=xlToLeft

xlApp.Range("O:O").Select
xlApp.Range("O:O").Activate
xlApp.Selection.NumberFormat = "0" '"#.0#"" per part"""

xlApp.Columns("O:O").Select
xlApp.Selection.Replace What:="""", Replacement:="", LookAt:=xlPart, _
SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
ReplaceFormat:=False

xlApp.Range("Q:Q").Select
xlApp.Range("Q:Q").Activate
xlApp.Selection.NumberFormat = "0.0000" '"#.0#"" per part"""

xlApp.Range("Q:Q").Select
xlApp.Range("Q:Q").Activate
xlApp.Selection.Replace What:="""", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
xlApp.Range("O:O,Q:Q").Select
xlApp.Range("O:O,Q:Q").Activate
With xlApp.Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
End With

'------> Activa los mensaje
xlApp.DisplayAlerts = True
xlWb.Close True, NomArchivoExcel

Dim XL As New excel.Application 'Crea el objeto excel
XL.Workbooks.Open NomArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
XL.Visible = True
XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada
    
'-------> Close ADO objects
RS.Close
Set RS = Nothing
    
' -- Cerrar Excel
xlApp.Quit
'-------> Release Excel references
Set xlWs = Nothing
Set xlWb = Nothing
Set xlApp = Nothing
  
fg_descarga
MsgBox "Proceso Finalizado", vbInformation, Me.Caption
                
Command1(0).Enabled = True
  
Exit Sub
Man_Error:
Command1(0).Enabled = True
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub ActualizarCambioTablaGramaje()

On Error GoTo Man_Error

Dim j               As Long
Dim dbexcel         As Database
Dim RS              As New ADODB.Recordset
Dim SheetName       As String
Dim IndColumna      As Long
Dim MyBuffer        As String
Dim MyBufferTotal   As String
Dim File_Ext        As String

Dim strArray()      As String
Dim intCount        As Integer
Dim UltRow          As Long

Dim Ceco            As String
Dim Regimen         As Long
Dim Receta          As Long
Dim IngOri          As String
Dim IngDes          As String
Dim CanDes          As Double

Dim lngRow          As Long
Dim NomArchivoExcel As String
Dim NomArchivo      As String

'Definición variables excel
Dim xlApp           As Object
Dim xlWb            As Object
Dim xlWs            As Object
   
Dim hoja            As String
Dim cs              As String
Dim sSheetName      As String
Dim PathXls         As String
Dim i               As Long

Dim RsExcel         As New ADODB.Recordset

Dim ApExcel         As New excel.Application
Set ApExcel = New excel.Application

NomArchivo = Dir(CD.FileName)

PathXls = Trim(fpText1.text)
hoja = Combo1.text '& "$"
   
RsExcel.CursorLocation = adUseClient
RsExcel.CursorType = adOpenKeyset
RsExcel.LockType = adLockBatchOptimistic
 
cs = "DRIVER=Microsoft Excel Driver (*.xls);DBQ=" & PathXls & ";HDR=NO;IMEX=1;"
' -- crea rnueva instancia de Excel

'Set obj_Excel = CreateObject("Excel.Application")
 
'obj_Excel.Visible = True

' -- Abrir el libro
ApExcel.Workbooks.Open FileName:=PathXls
'Set obj_Workbook = obj_Excel.Workbooks.Open(PathXls)
'obj_Excel.Visible = False
'obj_Excel.WindowState = xlMaximized 'Para que la ventana aparezca maximizada
' -- referencia la Hoja, por defecto la hoja activa

If sSheetName = vbNullString Then
   
'   Set obj_Worksheet = obj_Workbook.ActiveSheet
      
   For i = 1 To ApExcel.Sheets.count '.Sheets.count
         
       If hoja = ApExcel.Sheets(i).Name Then 'obj_Workbook.Sheets(i).Name Then
        
          hoja = ApExcel.Sheets(i).Name 'obj_Workbook.Sheets(i).Name
          Exit For
            
       End If
    
   Next
      
Else
'   Set obj_Worksheet = obj_Workbook.Sheets(hoja)
'
   hoja = ApExcel.Sheets(1).Name 'obj_Workbook.Sheets(hoja)
   
End If

hoja = "[" & hoja & "$" & "]"
RsExcel.Open "SELECT * FROM " & hoja, cs
 
i = 1

Let MyBuffer = ""
Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
Let MyBuffer = MyBuffer & "<UpdateGramaje>"

prbStatus.Max = 1
lblStatus.Visible = True
prbStatus.Visible = True
prbStatus.Min = 0
lngRow = 0
Frame1.Enabled = False
Command1(0).Enabled = False
prbStatus.Max = RsExcel.RecordCount

lblStatus.Caption = "Preparando datos para actualizar"

Dim XL As New excel.Application 'Crea el objeto excel

Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Prepara_Modificar"), Me.HelpContextID, "", "", "")

Do While RsExcel.EOF <> True

   If RsExcel.Fields(0).Value = "*" Then Exit Do

   Ceco = ""
   Regimen = 0
   Receta = 0
   IngOri = ""
   IngDes = ""
   CanDes = 0

   'Ceco
   If Not IsNull(RsExcel.Fields(0).Value) Then

      Ceco = RsExcel.Fields(0).Value

   End If

   'Regimen
   If IsNumeric(RsExcel.Fields(2).Value) Then

      Regimen = RsExcel.Fields(2).Value

   End If

   'Receta
   If IsNumeric(RsExcel.Fields(4).Value) Then

      Receta = RsExcel.Fields(4).Value

   End If

   'Ing. Origen
   If Not IsNull(RsExcel.Fields(11).Value) Then

      IngOri = RsExcel.Fields(11).Value

   End If

   'Ing. Destino
   If Not IsNull(RsExcel.Fields(14).Value) Then

      IngDes = RsExcel.Fields(14).Value

   End If

   'Can. Destino
   If IsNumeric(RsExcel.Fields(16).Value) Then

      CanDes = RsExcel.Fields(16).Value

   End If

   If Ceco <> "" And Regimen > 0 And Receta > 0 And IngOri <> "" And IngDes <> "" And CanDes >= 0 Then

      MyBuffer = MyBuffer & " <Gramaje"
      MyBuffer = MyBuffer & " Ceco = " & Chr(34) & Ceco & Chr(34)
      MyBuffer = MyBuffer & " Reg = " & Chr(34) & Regimen & Chr(34)
      MyBuffer = MyBuffer & " Rec = " & Chr(34) & Receta & Chr(34)
      MyBuffer = MyBuffer & " IOr = " & Chr(34) & IngOri & Chr(34)
      MyBuffer = MyBuffer & " IDe = " & Chr(34) & IngDes & Chr(34)
      MyBuffer = MyBuffer & " cDe = " & Chr(34) & CanDes & Chr(34)
      MyBuffer = MyBuffer & "/>"

   End If

   DoEvents

   RsExcel.MoveNext

   lngRow = lngRow + 1
   prbStatus.Value = lngRow

   If i > 1000 Then

      fg_carga ""
      lblStatus.Caption = "Actualizando tabla de gramaje ceco"

      MyBuffer = MyBuffer & "</UpdateGramaje>"

      Set RS = vg_db.Execute("sgpadm_UpdIns_XmlCambioTablaGramaje_V01 '" & MyBuffer & "', '" & UCase(LimpiaDato(Trim(vg_NUsr))) & "'")

      If Not RS.EOF Then

         If RS(0) > 0 Or RS(0) < 0 Then

            lblStatus.Visible = False
            prbStatus.Visible = False
            Frame1.Enabled = True
            Command1(0).Enabled = True
            fg_descarga

            Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Modificacion"), Me.HelpContextID, "", "", "")

            MsgBox RS(1) & VgLinea & VgLinea & "            Adjunto archivo con error", vbCritical, MsgTitulo

            '-------> Create an instance of Excel and add a workbook
            Set xlApp = CreateObject("Excel.Application")
            Set xlWb = xlApp.Workbooks.Add
            Set xlWs = xlWb.Worksheets("Hoja1")

            '-------> Display Excel and give user control of Excel's lifetime
            xlApp.UserControl = True

            '-------> Check version of Excel
            Call encabezado(RS, xlWs)

            xlWs.Cells(2, 1).CopyFromRecordset RS

            '-------> Auto-fit the column widths and row heights
            xlApp.Selection.CurrentRegion.Columns.AutoFit
            xlApp.Selection.CurrentRegion.Rows.AutoFit

            xlApp.Columns("A:B").Select
            xlApp.Selection.Delete Shift:=xlToLeft

            NomArchivoExcel = fg_ArchivoXls("ReporteError_actualizaciontablagramaje")

            xlWb.Close True, NomArchivoExcel

'            Dim XL As New excel.Application 'Crea el objeto excel
            XL.Workbooks.Open NomArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
            XL.Visible = True
            XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada

            '-- Cerrar Excel
            xlApp.Quit

            '-------> Release Excel references
            Set xlWs = Nothing
            Set xlWb = Nothing
            Set xlApp = Nothing

            RS.Close
            Set RS = Nothing

            
            '-- Cerrar aplicación excel
            ApExcel.Workbooks(NomArchivo).Close SaveChanges:=False
            
            ApExcel.Visible = False
            ApExcel.Application.Visible = False
            ApExcel.Application.Quit
            Set ApExcel = Nothing
            
            Exit Sub

         Else

            Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Modificado"), Me.HelpContextID, "", "", "")
            'MsgBox "Proceso finalizado sin problema", vbInformation + vbOKOnly, Me.Caption

         End If

      End If
      RS.Close
      Set RS = Nothing

      fg_descarga
      lblStatus.Caption = "Preparando datos para actualizar"
      i = 1

      Let MyBuffer = ""
      Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
      Let MyBuffer = MyBuffer & "<UpdateGramaje>"

   End If
   i = i + 1

Loop
   
'--Set Leer_Excel = RsExcel
Set RsExcel = Nothing
 
'' -- Cerrar libro
'obj_Workbook.Close

'' -- Cerrar Excel
'obj_Excel.Quit
'Set obj_Excel = Nothing
'obj_Workbook.Close ' SaveChanges:=False
'Set obj_Workbook = Nothing
'Set obj_Worksheet = Nothing
'
'Set obj_Workbook = Nothing
'Set obj_Excel = Nothing
'Set obj_Worksheet = Nothing
 
'-- Cerrar aplicación excel
ApExcel.Workbooks(NomArchivo).Close SaveChanges:=False

ApExcel.Visible = False
ApExcel.Application.Visible = False
ApExcel.Application.Quit
Set ApExcel = Nothing

MyBuffer = MyBuffer & "</UpdateGramaje>"

Set RS = vg_db.Execute("sgpadm_UpdIns_XmlCambioTablaGramaje_V01 '" & MyBuffer & "', '" & UCase(LimpiaDato(Trim(vg_NUsr))) & "'")

If Not RS.EOF Then

   If RS(0) > 0 Or RS(0) < 0 Then

      lblStatus.Visible = False
      prbStatus.Visible = False
      Frame1.Enabled = True
      Command1(0).Enabled = True
      fg_descarga

      Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Modificacion"), Me.HelpContextID, "", "", "")

      MsgBox RS(1) & VgLinea & VgLinea & "            Adjunto archivo con error", vbCritical, MsgTitulo

      '-------> Create an instance of Excel and add a workbook
      Set xlApp = CreateObject("Excel.Application")
      Set xlWb = xlApp.Workbooks.Add
      Set xlWs = xlWb.Worksheets("Hoja1")

      '-------> Display Excel and give user control of Excel's lifetime
      xlApp.UserControl = True

      '-------> Check version of Excel
      Call encabezado(RS, xlWs)

      xlWs.Cells(2, 1).CopyFromRecordset RS

      '-------> Auto-fit the column widths and row heights
      xlApp.Selection.CurrentRegion.Columns.AutoFit
      xlApp.Selection.CurrentRegion.Rows.AutoFit

      xlApp.Columns("A:B").Select
      xlApp.Selection.Delete Shift:=xlToLeft

      NomArchivoExcel = fg_ArchivoXls("ReporteError_actualizacióntablagramaje")

      xlWb.Close True, NomArchivoExcel

      XL.Workbooks.Open NomArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
      XL.Visible = True
      XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada

      '-- Cerrar Excel
      xlApp.Quit

      '-------> Release Excel references
      Set xlWs = Nothing
      Set xlWb = Nothing
      Set xlApp = Nothing

      RS.Close
      Set RS = Nothing

      Exit Sub

   Else

      Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Modificado"), Me.HelpContextID, "", "", "")

      MsgBox "Proceso finalizado sin problema", vbInformation + vbOKOnly, Me.Caption

   End If

End If
RS.Close
Set RS = Nothing

lblStatus.Visible = False
prbStatus.Visible = False
Frame1.Enabled = True
Command1(0).Enabled = True
fg_descarga

Exit Sub
Man_Error:
Command1(0).Enabled = True
Frame1.Enabled = True
fg_descarga

If Err = 5 Or Err = 462 Or Err = 430 Or Err = -2147023170 Or Err = -2147417848 Then Resume Next: Exit Sub
If Err = 55 Then MsgBox "Archivo esta abierto, proceso cancelado": Exit Sub
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Command2_Click(Index As Integer)

On Error GoTo Man_Error

B_DieTipExcel.Show 1

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Command3_Click()

On Error GoTo Man_Error

vg_left = Command3.Left + 10300
Call B_TablaEstandar.Show(1)

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Activate()

On Error GoTo Man_Error

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

fg_carga ""
Me.HelpContextID = vg_OpcM

MsgTitulo = "Exportar Excel tabla Gramaje & Bach - Input"
fg_centra Me

Command1(0).Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1) = "0", False, True)

fpText1.text = ""
fpText1.Enabled = False
Combo1.Enabled = False
Combo1.ListIndex = -1

lblStatus.Visible = False
prbStatus.Visible = False

FilCatDie = 0
FilTipPla = 0

B_DieTipExcel.MoverDatosTvwDir
B_TablaEstandar.CargaMaestros 1

fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpLongInteger1_Change(Index As Integer)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Select Case Index

Case 1
    
    If vg_Indppr = 1 Or vg_Indppr = 2 Then
      
      Set RS = vg_db.Execute("SELECT * FROM a_regimen WHERE reg_codigo = " & Val(fpLongInteger1(1).Value) & " AND reg_indppr = '" & vg_Indppr & "'")
    
    Else
      
      Set RS = vg_db.Execute("SELECT * FROM a_regimen WHERE reg_codigo = " & Val(fpLongInteger1(1).Value) & "")
    
    End If
    
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(4).Caption = "": Exit Sub
    fpayuda(4).Caption = Trim(RS!reg_nombre)
    RS.Close: Set RS = Nothing
  
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText_Change(Index As Integer)
On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim Sql As String

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Select Case Index

Case 1
   
   Sql = Trim(LimpiaDato(fpText(1).text))
   Set RS = vg_db.Execute("sgpadm_s_cliente_V02 29, '" & Sql & "', ''")
   If RS.EOF Then
        
        fpayuda(9).Caption = ""
        RS.Close
        Set RS = Nothing
        Exit Sub
    
    End If
    fpayuda(9).Caption = Trim(RS!Cli_nombre)

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText1_ButtonHit(Button As Integer, NewIndex As Integer)

On Error GoTo Man_Error

Dim fromRihgt  As String
Dim myPath     As String
Dim NomArchivo As String

CD.Filter = "Archivos xls|*.xls|Archivos xlsx|*.xlsx"
CD.DefaultExt = "*.xls|*.xlsx"
CD.FilterIndex = 2
CD.Flags = cdlOFNFileMustExist
CD.Flags = &H80000 Or &H400& Or &H1000& Or &H4&
CD.ShowOpen

If CD.FileName = "" Then
   
   fpText1.text = ""

Else

    Combo1.Clear
    
    fpText1.text = CD.FileName 'Dir(CD.FileName)

    
'    Dim ObjExcel As excel.Application
'    Dim ObjW As excel.Workbook
'
'    Set ObjExcel = New excel.Application
'    Set ObjW = ObjExcel.Workbooks.Open(fpText1.text)
'    Dim i As Integer
'    Dim HojaPro As Boolean
'
'    For i = 1 To ObjW.Sheets.count
'
'        Combo1.AddItem ObjW.Sheets(i).Name
'    Next
'
'    ObjW.Application.DisplayAlerts = False
'    ObjW.Close
'    Set ObjExcel = Nothing
'    Set ObjW = Nothing
    
    NomArchivo = Dir(CD.FileName)

    Dim ApExcel As excel.Application

    'Al configurarlo
    Set ApExcel = New excel.Application

    'Al abrirlo
    ApExcel.Workbooks.Open FileName:=fpText1.text
    
    Dim i As Long

    For i = 1 To ApExcel.Sheets.count

        Combo1.AddItem ApExcel.Sheets(i).Name
        
    Next

    ApExcel.Workbooks(NomArchivo).Close SaveChanges:=False

    ApExcel.Visible = False
    ApExcel.Application.Visible = False
    ApExcel.Application.Quit
    Set ApExcel = Nothing

End If

Exit Sub
Man_Error:
fg_descarga
If Err = 5 Then Resume Next: Exit Sub
If Err = 55 Then MsgBox "Archivo esta abierto, proceso cancelado": Exit Sub
If Err = 462 Or Err = 1004 Or Err = 438 Then Resume Next
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
Resume

End Sub

Private Sub Image1_Click(Index As Integer)


On Error GoTo Man_Error

Select Case Index

Case 4
    
    vg_left = fpayuda(4).Left + 3000
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_regimen", "reg_", "Regimen", "Reg"
    B_TabEst.Show 1
    Me.Refresh
    Screen.MousePointer = 0
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(1).Value = Val(vg_codigo)
    fpayuda(4).Caption = vg_nombre
'    fpLongInteger1(2).SetFocus
 
 Case 6
    
    vg_left = fpayuda(9).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Casino", "CliAct"
    B_TabEst.Show 1
    Me.Refresh
    Screen.MousePointer = 0
    If vg_codigo = "" Then Exit Sub
    fpText(1).text = vg_codigo
    fpayuda(9).Caption = vg_nombre
    fpLongInteger1(1).SetFocus

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Option1_Click(Index As Integer)

On Error GoTo Man_Error

Select Case Index

    Case 0
    
        fpText1.text = ""
        
        fpText1.Enabled = False
        Combo1.Enabled = False
        Combo1.ListIndex = -1
    
        fpText(1).text = ""
        fpText(1).Enabled = True
        Image1(6).Enabled = True
        fpayuda(9).Caption = ""
        fpLongInteger1(1).text = ""
        fpLongInteger1(1).Enabled = True
        Image1(4).Enabled = True
        fpayuda(4).Caption = ""
        Command2(0).Enabled = True
        Command3.Enabled = True
        Frame2.Enabled = True

    Case 1
    
        fpText(1).text = ""
        fpText(1).Enabled = False
        Image1(6).Enabled = False
        fpayuda(9).Caption = ""
        fpLongInteger1(1).text = ""
        fpLongInteger1(1).Enabled = False
        Image1(4).Enabled = False
        fpayuda(4).Caption = ""
        Command2(0).Enabled = False
        Command3.Enabled = False
        Frame2.Enabled = False
        
        fpText1.Enabled = True
        fpText1.text = ""
        Combo1.Enabled = True
        Combo1.ListIndex = -1

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Function ValidaDatos() As Boolean

On Error GoTo Man_Error

Dim SheetName As String
'Dim cn        As New ADODB.Connection
Dim RsExcel   As New ADODB.Recordset
Dim PathXls   As String
Dim File_Ext  As String
Dim dbexcel   As Database
Dim i         As Long
'Set cn = New ADODB.Connection

Let ValidaDatos = True
 
If Option1(1).Value = True Then

    Set RsExcel = New ADODB.Recordset
    Dim hoja As String
    Dim cs As String
    Dim sSheetName As String
   
    PathXls = Trim(LimpiaDato(fpText1.text))
    '-------> Validar Archivo Origen
    If Trim(LimpiaDato(fpText1.text)) = "" Then
        
        Call MsgBox("Debe seleccionar archivo origen", vbInformation, Me.Caption)
        Call fpText1.SetFocus
        Let ValidaDatos = False
        Exit Function
    
    End If
    
    '-------> Validar hoja
    If Combo1.ListIndex = -1 Then
        
        Call MsgBox("Debe seleccionar hoja", vbInformation, Me.Caption)
        Call Combo1.SetFocus
        Let ValidaDatos = False
        Exit Function
    
    End If
    
   RsExcel.CursorLocation = adUseClient
   RsExcel.CursorType = adOpenKeyset
   RsExcel.LockType = adLockBatchOptimistic
 
   cs = "DRIVER=Microsoft Excel Driver (*.xls);DBQ=" & PathXls & ";HDR=NO;IMEX=1;"
   ' -- crea rnueva instancia de Excel

   Set obj_Excel = CreateObject("Excel.Application")
 
   'obj_Excel.Visible = True

 
   ' -- Abrir el libro
   Set obj_Workbook = obj_Excel.Workbooks.Open(PathXls)
   obj_Excel.Visible = False
   ' -- referencia la Hoja, por defecto la hoja activa

   hoja = Combo1.text
   If sSheetName = vbNullString Then
      Set obj_Worksheet = obj_Workbook.ActiveSheet
      
      For i = 1 To obj_Workbook.Sheets.count
 
          If hoja = obj_Workbook.Sheets(i).Name Then
        
             hoja = obj_Workbook.Sheets(i).Name
             Exit For
            
          End If
    
      Next
      
   Else
      Set obj_Worksheet = obj_Workbook.Sheets(hoja)
'
      hoja = obj_Workbook.Sheets(hoja)
   
   End If
 
   hoja = "[" & hoja & "$" & "]"
   RsExcel.Open "SELECT * FROM " & hoja, cs
   
   If RsExcel.EOF Then
       
       Call MsgBox("No existe información", vbInformation, Me.Caption)
       Let ValidaDatos = False
       RsExcel.Close
       Set RsExcel = Nothing
       obj_Workbook.Close
       obj_Excel.Quit
       Set obj_Workbook = Nothing
       Set obj_Excel = Nothing
       Set obj_Worksheet = Nothing
       
       Exit Function
       
    End If
    
    RsExcel.MoveFirst
    
    If RsExcel.Fields(0).Value = "*" Then
       
       Call MsgBox("Formato no corresponde", vbInformation, Me.Caption)
       Let ValidaDatos = False
       RsExcel.Close
       Set RsExcel = Nothing
       obj_Workbook.Close
       obj_Excel.Quit
       Set obj_Workbook = Nothing
       Set obj_Excel = Nothing
       Set obj_Worksheet = Nothing
       
       Exit Function
       
    End If
    
    'Ceco
    If Not Trim(RsExcel.Fields(0).Name) = "Ceco" Then
    
       Call MsgBox("Formato Ceco no corresponde", vbInformation, Me.Caption)
       Let ValidaDatos = False
       RsExcel.Close
       Set RsExcel = Nothing
       obj_Workbook.Close
       obj_Excel.Quit
       Set obj_Workbook = Nothing
       Set obj_Excel = Nothing
       Set obj_Worksheet = Nothing
       
       Exit Function
    
    End If
    
    'Nombre servicio
    If Not Trim(RsExcel.Fields(2).Name) = "Regimen" Then
       
       Call MsgBox("Formato Régimen no corresponde", vbInformation, Me.Caption)
       Let ValidaDatos = False
       RsExcel.Close
       Set RsExcel = Nothing
       obj_Workbook.Close
       obj_Excel.Quit
       Set obj_Workbook = Nothing
       Set obj_Excel = Nothing
       Set obj_Worksheet = Nothing
       
       Exit Function
       
    End If
       
    'Receta
    If Not Trim(RsExcel.Fields(4).Name) = "Receta" Then
       
       Call MsgBox("Formato Receta no corresponde", vbInformation, Me.Caption)
       Let ValidaDatos = False
       RsExcel.Close
       Set RsExcel = Nothing
       obj_Workbook.Close
       obj_Excel.Quit
       Set obj_Workbook = Nothing
       Set obj_Excel = Nothing
       Set obj_Worksheet = Nothing
       
       Exit Function
       
    End If
       
    'Ing. Origen
    If Not Trim(RsExcel.Fields(11).Name) = "Ing# Origen" Then
       
       Call MsgBox("Formato Ing. Origen no corresponde", vbInformation, Me.Caption)
       Let ValidaDatos = False
       RsExcel.Close
       Set RsExcel = Nothing
       obj_Workbook.Close
       obj_Excel.Quit
       Set obj_Workbook = Nothing
       Set obj_Excel = Nothing
       Set obj_Worksheet = Nothing
       
       Exit Function
       
    End If
       
    'Cant Origen
    If Not Trim(RsExcel.Fields(13).Name) = "Cant Origen" Then
       
       Call MsgBox("Formato Cant. Origen no corresponde", vbInformation, Me.Caption)
       Let ValidaDatos = False
       RsExcel.Close
       Set RsExcel = Nothing
       obj_Workbook.Close
       obj_Excel.Quit
       Set obj_Workbook = Nothing
       Set obj_Excel = Nothing
       Set obj_Worksheet = Nothing
       
       Exit Function
       
    End If
       
    'Ing# Alt#
    If Not Trim(RsExcel.Fields(14).Name) = "Ing# Alt#" Then
       
       Call MsgBox("Formato Ing. Alt. no corresponde", vbInformation, Me.Caption)
       Let ValidaDatos = False
       RsExcel.Close
       Set RsExcel = Nothing
       obj_Workbook.Close
       obj_Excel.Quit
       Set obj_Workbook = Nothing
       Set obj_Excel = Nothing
       Set obj_Worksheet = Nothing
       
       Exit Function
       
    End If
       
    'Cant. Alt.
    If Not Trim(RsExcel.Fields(16).Name) = "Cant# Alt#" Then
       
       Call MsgBox("Formato Cant. Alt. no corresponde", vbInformation, Me.Caption)
       Let ValidaDatos = False
       RsExcel.Close
       Set RsExcel = Nothing
       obj_Workbook.Close
       obj_Excel.Quit
       Set obj_Workbook = Nothing
       Set obj_Excel = Nothing
       Set obj_Worksheet = Nothing
       
       Exit Function
       
    End If
    
    DoEvents
               
      RsExcel.Close
      Set RsExcel = Nothing
      obj_Workbook.Close
      obj_Excel.Quit
      Set obj_Workbook = Nothing
      Set obj_Excel = Nothing
      Set obj_Worksheet = Nothing

ElseIf Option1(0).Value = True Then

    If fpayuda(9).Caption = "" Then

       Call MsgBox("Debe seleccionar Ceco", vbInformation, Me.Caption)
       Call fpText(1).SetFocus
       Let ValidaDatos = False
       Exit Function
    
    End If

    If Trim(fpLongInteger1(1).text) <> "" Then

       If fpayuda(4).Caption = "" Then
       
          Call MsgBox("Debe seleccionar Régimen", vbInformation, Me.Caption)
          Call fpLongInteger1(1).SetFocus
          Let ValidaDatos = False
          Exit Function
          
       End If
       
    End If

End If
    
Exit Function
Man_Error:
    ValidaDatos = False
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo
  
End Function
