VERSION 5.00
Begin VB.Form M_BorrarArchivos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantendor Borrado Archivos de Sistema"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   3045
      TabIndex        =   1
      Top             =   315
      Width           =   2220
   End
   Begin VB.FileListBox File1 
      Height          =   2625
      Left            =   210
      TabIndex        =   0
      Top             =   315
      Width           =   2220
   End
   Begin VB.Label lblcant 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   210
      TabIndex        =   5
      Top             =   3045
      Width           =   855
   End
   Begin VB.Label lblcant_B 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3045
      TabIndex        =   4
      Top             =   3045
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Lista de archivos por borrar"
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
      Left            =   3045
      TabIndex        =   3
      Top             =   105
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Universo de Archivos a borrar"
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
      Left            =   210
      TabIndex        =   2
      Top             =   105
      Width           =   2595
   End
End
Attribute VB_Name = "M_BorrarArchivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub InicioBorrado(Path As String, NomFile As String, NumArc As Long)
Dim i As Long
Dim z As Long
On Error GoTo e
'-------> Esta rutina va mantendor los cinco ultimo backup del sistema
z = NumArc '4
Me.File1.Path = Path
'-------> Filtro
Me.File1.Pattern = NomFile '"dbgt_alemana*.zip"

lblcant.Caption = Me.File1.listcount

Me.List1.Clear
For i = 0 To Me.File1.listcount - (z + 1)
    Me.List1.AddItem (File1.List(i))
    If Dir(Path & File1.List(i)) <> "" Then Kill Path & File1.List(i)
Next i
Me.lblcant_B.Caption = Me.List1.listcount

Exit Sub
e:
    MsgBox Err.Description, vbCritical
    
End Sub

