VERSION 5.00
Begin VB.Form Busqueda 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Buscando..."
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5730
   ControlBox      =   0   'False
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox listDirs 
      Height          =   1425
      Left            =   525
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.DirListBox dirTemp 
      Height          =   540
      Left            =   150
      TabIndex        =   0
      Top             =   900
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5520
   End
End
Attribute VB_Name = "Busqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 Me.MousePointer = 11
End Sub

Public Sub Recopilar(RutaInicial As String, ListaArchivo As Object, ListaDir As Object, Exten As String, Mensaje As String, UsarDire As Boolean)
 Dim txtfile As String
 Dim Y As Integer
 Dim path As String
 Dim tfilename As String
    Me.Caption = SPELL(Mensaje)
    ListaArchivo.Clear
    If UsarDire = True Then ListaDir.Clear
    listDirs.Clear
    listDirs.AddItem RutaInicial
    Y = 0
    Do Until Y = listDirs.ListCount
     DoEvents
       dirTemp.path = listDirs.List(Y)
        If dirTemp.ListCount > 0 Then
           For x = 0 To dirTemp.ListCount - 1
            listDirs.AddItem dirTemp.List(x)
           Next x
        End If
        Y = Y + 1
    Loop
    For x = 0 To listDirs.ListCount - 1
     DoEvents
     If listDirs.List(x) Like "*\" Then
      txtfile = Dir(listDirs.List(x) & "*." + Trim(Exten))
     Else
      txtfile = Dir(listDirs.List(x) & "\*." + Trim(Exten))
     End If
     Label1.Caption = SPELL(listDirs.List(x))
     If Not txtfile = "" Then
        Do
            tfilename = path & txtfile
            ListaArchivo.AddItem tfilename
            If UsarDire = True Then
             If Len(listDirs.List(x)) = 3 Then ListaDir.AddItem DepurarUbi(listDirs.List(x))
             If Len(listDirs.List(x)) > 3 Then ListaDir.AddItem DepurarUbi(listDirs.List(x) & "\")
            End If
            txtfile = Dir$
        Loop Until txtfile = ""
     End If
    Next x
End Sub
