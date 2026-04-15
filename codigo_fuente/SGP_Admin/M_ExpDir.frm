VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_ExpDir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Explorador Directorio"
   ClientHeight    =   4575
   ClientLeft      =   7365
   ClientTop       =   3285
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   1020
      TabIndex        =   2
      Top             =   4110
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2460
      TabIndex        =   1
      Top             =   4110
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   60
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_ExpDir.frx":0000
            Key             =   "Computer"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_ExpDir.frx":0454
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_ExpDir.frx":08A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_ExpDir.frx":0CFC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwDir 
      Height          =   3375
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5953
      _Version        =   393217
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.Label lblPath 
      Caption         =   "lblPath"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "M_ExpDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dr As Scripting.Drive
' True si se ha pulsado Cancelar para cerrar este formulario
Public CancelPressed As Boolean
Private m_Path As String
' lo siguiente lo utilizan algunas rutinas del módulo
Dim fso As New Scripting.FileSystemObject

' la Ruta (Path) seleccionada actualmente
Property Get Path() As String
    Path = m_Path
End Property

Private Sub Form_Load()

' construir el árbol de subdirectorios
    DirRefresh

End Sub

Private Sub Form_Resize()
    
    ' distancia entre controles
    Const DISTANCE = 100
    Dim tvwTop As Single
    
    ' mover los botones y el rótulo
    lblPath.Move DISTANCE, 0, ScaleWidth, lblPath.Height
    cmdOK.Move ScaleWidth / 2 - DISTANCE - cmdOK.Width, ScaleHeight - DISTANCE - cmdOK.Height
    cmdCancel.Move ScaleWidth / 2 + DISTANCE, cmdOK.Top
    ' cambiar el tamaño al control treeview
    ' la posición Top depende de la visibilidad del rótulo lblPath
    If lblPath.Visible Then
        tvwTop = lblPath.Top + lblPath.Height
    Else
        tvwTop = DISTANCE
    End If
    TvwDir.Move DISTANCE, tvwTop, ScaleWidth - DISTANCE * 2, ScaleHeight - tvwTop - cmdOK.Height - DISTANCE * 2

End Sub

Private Sub DirRefresh()
    
    ' generar el control treeview
    On Error Resume Next
    
    Dim rootNode As Node, nd As Node
    ' agregar la raíz "Mi PC" (expandida)
    Set rootNode = TvwDir.Nodes.Add(, , "\\Mi PC", "Mi PC", 1)
    rootNode.Expanded = True
    ' agregar todas las unidades, con un signo más
    For Each dr In fso.Drives
'        If dr.Path <> "A:" Then
        Err.Clear
        If dr.IsReady Then
           Set nd = TvwDir.Nodes.Add(rootNode.key, tvwChild, dr.Path & "\", dr.Path & " " & dr.VolumeName, 2)
           If Err = 0 Then AddDummyChild nd
        Else
           Set nd = TvwDir.Nodes.Add(rootNode.key, tvwChild, dr.Path & "\", dr.Path & " ", 2)
'           If Err = 0 Then AddDummyChild nd
        End If
'        End If
    Next

End Sub

Sub AddDummyChild(nd As Node)

' agregar un nodo hijo postizo, si fuera necesario
If nd.Children = 0 Then
   
   ' la propiedad Texto de los nodos postizos es "***"
   TvwDir.Nodes.Add nd.Index, tvwChild, , "***"

End If

End Sub

Private Sub TvwDir_Click()

m_Path = TvwDir.SelectedItem.key
lblPath.Caption = TvwDir.SelectedItem.key

End Sub

Private Sub tvwDir_Expand(ByVal Node As MSComctlLib.Node)

' se ha expandido un nodo
Dim nd As Node
' salir si el nodo ya fue expandido en el pasado
If Node.Children = 0 Or Node.Children > 1 Then Exit Sub
' salir también si no existe un nodo hijo postizo
If Node.Child.text <> "***" Then Exit Sub
' eliminar el elemento hijo postizo
TvwDir.Nodes.Remove Node.Child.Index
' agregar todos los subdirectorios de este objeto Nodo
AddSubdirs Node

End Sub

Private Sub AddSubdirs(ByVal Node As MSComctlLib.Node)

On Error Resume Next
' agregar todos los subdirectorios de este objeto Nodo
Dim fld As Scripting.Folder
Dim nd As Node

' la ruta en el nodo está almacenada en su propiedad clave, por lo que resulta sencillo
' hacer un ciclo en todos sus subdirectorios
For Each fld In fso.GetFolder(Node.key).SubFolders
    Set nd = TvwDir.Nodes.Add(Node, tvwChild, fld.Path, fld.Name, 3)
    nd.ExpandedImage = 4
    ' si este subdirectorio cuenta con subcarpetas, agregar un signo "+"
    If fld.SubFolders.count Then AddDummyChild nd
Next

End Sub

Private Sub cmdOK_Click()
    
    vg_dir = lblPath.Caption
    Unload Me

End Sub

Private Sub cmdCancel_Click()
    
    CancelPressed = True
    Unload Me

End Sub
