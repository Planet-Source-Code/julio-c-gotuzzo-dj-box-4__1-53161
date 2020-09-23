VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Nueva Lista De Reproducción"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8070
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   3960
      Index           =   1
      Left            =   5700
      TabIndex        =   7
      Top             =   900
      Width           =   1290
   End
   Begin VB.ListBox List2 
      Height          =   3960
      Index           =   0
      Left            =   2250
      TabIndex        =   6
      Top             =   1425
      Width           =   1290
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<<"
      Height          =   315
      Index           =   3
      Left            =   3825
      TabIndex        =   5
      Top             =   3450
      Width           =   390
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      Height          =   315
      Index           =   2
      Left            =   3825
      TabIndex        =   4
      Top             =   3075
      Width           =   390
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   315
      Index           =   1
      Left            =   3825
      TabIndex        =   3
      Top             =   2700
      Width           =   390
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">>"
      Height          =   315
      Index           =   0
      Left            =   3825
      TabIndex        =   2
      Top             =   2325
      Width           =   390
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   6105
      Index           =   1
      Left            =   4275
      TabIndex        =   1
      Top             =   75
      Width           =   3690
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   6105
      Index           =   0
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   3690
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 Dim n As Integer
 n = 0
 Do While n <= Form1.LISTA(0).ListCount - 1
  List1(0).AddItem Form1.LISTA(0).List(n)
  List2(0).AddItem Form1.LISTA(1).List(n)
  n = n + 1
 Loop
End Sub
