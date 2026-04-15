VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5565
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8010
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   8010
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   4155
      Left            =   4320
      TabIndex        =   1
      Top             =   600
      Width           =   3495
   End
   Begin VB.FileListBox File1 
      Height          =   4185
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3855
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
      Left            =   4320
      TabIndex        =   5
      Top             =   120
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
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label lblcant_B 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label lblcant 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5040
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim i As Long
Dim z As Long
On Error GoTo e
z = 5
Me.File1.Path = "C:\Archivos de programa\sgp\Backup\"
'Filtro
Me.File1.Pattern = "dbgt_alemana*.zip"

lblcant.Caption = Me.File1.ListCount

Me.List1.Clear
For i = 0 To Me.File1.ListCount - (z + 1)
    Me.List1.AddItem (File1.List(i))
Next i
Me.lblcant_B.Caption = Me.List1.ListCount
Exit Sub
e:
    MsgBox Err.Description, vbCritical
End Sub
