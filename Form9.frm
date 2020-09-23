VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Seleccione o cree una lista de reproducción"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4500
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   1650
      Left            =   3075
      Pattern         =   "*.m3u;*.pls"
      TabIndex        =   5
      Top             =   1875
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   1620
      Index           =   1
      Left            =   -750
      TabIndex        =   4
      Top             =   1950
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   240
      Index           =   1
      Left            =   3075
      TabIndex        =   3
      Top             =   5175
      Width           =   1290
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Crear Nueva"
      Height          =   240
      Index           =   0
      Left            =   75
      TabIndex        =   2
      Top             =   5175
      Width           =   1290
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   4155
      Index           =   0
      ItemData        =   "Form9.frx":0000
      Left            =   75
      List            =   "Form9.frx":0002
      TabIndex        =   1
      Top             =   525
      Width           =   4290
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   75
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   75
      Width           =   4290
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tiempo De Reproducción: 00:00:00"
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   75
      TabIndex        =   6
      Top             =   4800
      Width           =   4290
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
 List1(0).Clear
 List1(1).Clear
 Call M3ULOAD("c:\sanson\dj box 4\listas\" + File1.List(Combo1.ListIndex), List1(0), List1(1))
 Call ordenar2(List1(0), List1(1), Indicador, False)
 Label1 = "Tiempo De Reproducción: " + HORMINSEG(temasdur(List1(1)))
End Sub

Private Sub Command1_Click(Index As Integer)
 Form11.Enabled = True
 Form11.Visible = True
End Sub

Private Sub Form_Load()
 Dim n As Integer
 File1.path = "c:\sanson\dj box 4\listas"
 n = 0
 Do While n <= File1.ListCount - 1
  Combo1.AddItem SPELL(SINEXT(File1.List(n)))
  n = n + 1
 Loop
 If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
End Sub
