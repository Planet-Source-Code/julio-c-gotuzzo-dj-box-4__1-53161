VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3570
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   4785
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "Frmspl~1.frx":0000
   ScaleHeight     =   3570
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   3825
      Top             =   3075
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Click()
 frmSplash.MousePointer = 0
 Form1.Enabled = True
 Form1.Visible = True
 Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 frmSplash.MousePointer = 0
 Form1.Enabled = True
 Form1.Visible = True
 Unload Me
End Sub

Private Sub Form_Load()
 frmSplash.MousePointer = 11
End Sub

Private Sub Timer1_Timer()
 frmSplash.MousePointer = 0
 Form1.Enabled = True
 Form1.Visible = True
 Unload Me
End Sub
