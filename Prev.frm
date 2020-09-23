VERSION 5.00
Begin VB.Form Prev 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Preview..."
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5805
   ControlBox      =   0   'False
   Icon            =   "Prev.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   255
      Left            =   4575
      TabIndex        =   1
      Top             =   6330
      Width           =   1095
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00400000&
      ForeColor       =   &H000080FF&
      Height          =   6105
      ItemData        =   "Prev.frx":0442
      Left            =   120
      List            =   "Prev.frx":0444
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "Prev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Unload Prev
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Unload Prev
End Sub
