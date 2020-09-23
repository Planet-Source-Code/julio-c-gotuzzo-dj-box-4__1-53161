VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Dj Box 4 - Cargar Fx"
   ClientHeight    =   3180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6735
   ControlBox      =   0   'False
   Icon            =   "fxrtm.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3360
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   -120
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2520
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   -120
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   90
      MaxLength       =   6
      TabIndex        =   3
      Top             =   2760
      Width           =   3180
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   2130
      Width           =   3150
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   1890
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3150
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   2430
      Left            =   3435
      Pattern         =   "*.mp3;*.wav"
      TabIndex        =   0
      Top             =   120
      Width           =   3210
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3645
      TabIndex        =   6
      Top             =   2805
      Width           =   1155
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5175
      TabIndex        =   5
      Top             =   2805
      Width           =   1155
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   90
      TabIndex        =   4
      Top             =   2535
      Width           =   1485
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   3240
      Left            =   -15
      Picture         =   "fxrtm.frx":0442
      Top             =   -30
      Width           =   6765
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim N As Integer

Private Sub Dir1_Change()
File1.path = Dir1.path
End Sub

Private Sub Drive1_Change()
 Dir1.path = Drive1.Drive
End Sub

Private Sub File1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Text1 = SINEXT(LCase(File1))
End Sub

Private Sub File1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Label2.ForeColor <> UNSELCOLOR Then Label2.ForeColor = UNSELCOLOR
 If Label1.ForeColor <> UNSELCOLOR Then Label1.ForeColor = UNSELCOLOR
End Sub

Private Sub Form_Load()
 If SKINFILE > 0 Then Call SET_SKIN(5, SKINFILE)
 Dir1.ForeColor = DATAFORE
 Dir1.BackColor = DATABACK
 Drive1.ForeColor = DATAFORE
 Drive1.BackColor = DATABACK
 File1.ForeColor = DATAFORE
 File1.BackColor = DATABACK
 Label1.ForeColor = UNSELCOLOR
 Label2.ForeColor = UNSELCOLOR
 Text1.ForeColor = DATAFORE
 Text1.BackColor = DATABACK
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Label2.ForeColor <> UNSELCOLOR Then Label2.ForeColor = UNSELCOLOR
 If Label1.ForeColor <> UNSELCOLOR Then Label1.ForeColor = UNSELCOLOR
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If File1.ListIndex > -1 Then
  Form3.MousePointer = 11
  If Text2 = 0 Then
   Form4.FX(Text3).Caption = LCase(Text1.Text)
   If Len(Dir1.path) = 3 Then Form4.EFFECT.List(Text3) = Dir1.path + File1
   If Len(Dir1.path) > 3 Then Form4.EFFECT.List(Text3) = Dir1.path + "\" + File1
  Else
   Form4.RTM(Text3).Caption = LCase(Text1.Text)
   If Len(Dir1.path) = 3 Then Form4.RITMO.List(Text3) = Dir1.path + File1
   If Len(Dir1.path) > 3 Then Form4.RITMO.List(Text3) = Dir1.path + "\" + File1
  End If
  File1.Selected(File1.ListIndex) = False
  Form3.MousePointer = 0
  Form1.Enabled = True
  Form4.Enabled = True
  Form5.Enabled = True
  Unload Me
 Else
  MsgBox ("Debe Seleccionar Un Archivo De Audio")
 End If
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Label1.ForeColor <> SELCOLOR Then Label1.ForeColor = SELCOLOR
 If Label2.ForeColor <> UNSELCOLOR Then Label2.ForeColor = UNSELCOLOR
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If File1.ListIndex > -1 Then File1.Selected(File1.ListIndex) = False
 Form1.Enabled = True
 Form4.Enabled = True
 Form5.Enabled = True
 Unload Me
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Label2.ForeColor <> SELCOLOR Then Label2.ForeColor = SELCOLOR
 If Label1.ForeColor <> UNSELCOLOR Then Label1.ForeColor = UNSELCOLOR
End Sub

Private Sub Dir1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Label2.ForeColor <> UNSELCOLOR Then Label2.ForeColor = UNSELCOLOR
 If Label1.ForeColor <> UNSELCOLOR Then Label1.ForeColor = UNSELCOLOR
End Sub
