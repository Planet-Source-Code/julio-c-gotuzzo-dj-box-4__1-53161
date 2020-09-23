VERSION 5.00
Begin VB.Form Form12 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "capo"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6810
   ControlBox      =   0   'False
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   240
      Index           =   1
      Left            =   5700
      TabIndex        =   8
      Top             =   5400
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   240
      Index           =   0
      Left            =   4725
      TabIndex        =   7
      Top             =   5400
      Width           =   990
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      Index           =   2
      Left            =   900
      TabIndex        =   6
      Top             =   900
      Width           =   5790
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      Index           =   1
      Left            =   900
      TabIndex        =   3
      Top             =   375
      Width           =   5790
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      Index           =   0
      Left            =   900
      TabIndex        =   1
      Top             =   75
      Width           =   5790
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   3960
      Left            =   75
      TabIndex        =   0
      Top             =   1350
      Width           =   6615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Pistas:"
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   75
      TabIndex        =   5
      Top             =   975
      Width           =   840
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Intérprete:"
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   75
      TabIndex        =   4
      Top             =   450
      Width           =   765
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Título:"
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   75
      TabIndex        =   2
      Top             =   150
      Width           =   540
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
 If Index = 0 Then
  
 End If
 Unload Form12
End Sub

Private Sub Form_Load()
 Form12.Caption = "Id. " + Trim(str(GetCDID))
 If LoadCDBase(List1, GetCDID, Text1(0), Text1(1), 2) = False Then
   Do While List1.ListCount < GetNumberOfTracks
    List1.AddItem "Pista " + str(List1.ListCount + 1)
   Loop
 End If
 List1.ListIndex = 0
End Sub

Private Sub Label3_Click()

End Sub

Private Sub List1_Click()
 Text1(2) = List1.List(List1.ListIndex)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
 If Index = 2 And KeyAscii = 13 Then List1.List(List1.ListIndex) = Text1(2)
End Sub
