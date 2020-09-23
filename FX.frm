VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Dj Box 4 - Fx Console"
   ClientHeight    =   3975
   ClientLeft      =   7755
   ClientTop       =   300
   ClientWidth     =   4095
   Icon            =   "FX.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox RITMO 
      Height          =   645
      ItemData        =   "FX.frx":0442
      Left            =   2160
      List            =   "FX.frx":0476
      TabIndex        =   1
      Top             =   3720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox EFFECT 
      Height          =   645
      ItemData        =   "FX.frx":04AA
      Left            =   720
      List            =   "FX.frx":04DE
      TabIndex        =   0
      Top             =   3720
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3780
      TabIndex        =   36
      ToolTipText     =   "Minimizar"
      Top             =   45
      Width           =   255
   End
   Begin VB.Label FX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vacío"
      Height          =   240
      Index           =   0
      Left            =   540
      TabIndex        =   35
      Top             =   585
      Width           =   675
   End
   Begin VB.Label FX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vacío"
      Height          =   240
      Index           =   1
      Left            =   1350
      TabIndex        =   34
      Top             =   585
      Width           =   675
   End
   Begin VB.Label FX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vacío"
      Height          =   240
      Index           =   2
      Left            =   2130
      TabIndex        =   33
      Top             =   585
      Width           =   675
   End
   Begin VB.Label FX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vacío"
      Height          =   240
      Index           =   3
      Left            =   2910
      TabIndex        =   32
      Top             =   585
      Width           =   675
   End
   Begin VB.Label FX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vacío"
      Height          =   240
      Index           =   4
      Left            =   540
      TabIndex        =   31
      Top             =   930
      Width           =   675
   End
   Begin VB.Label FX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vacío"
      Height          =   240
      Index           =   5
      Left            =   1350
      TabIndex        =   30
      Top             =   930
      Width           =   675
   End
   Begin VB.Label FX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vacío"
      Height          =   240
      Index           =   6
      Left            =   2130
      TabIndex        =   29
      Top             =   930
      Width           =   675
   End
   Begin VB.Label FX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vacío"
      Height          =   240
      Index           =   7
      Left            =   2910
      TabIndex        =   28
      Top             =   930
      Width           =   675
   End
   Begin VB.Label FX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vacío"
      Height          =   240
      Index           =   8
      Left            =   540
      TabIndex        =   27
      Top             =   1275
      Width           =   675
   End
   Begin VB.Label FX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vacío"
      Height          =   240
      Index           =   9
      Left            =   1350
      TabIndex        =   26
      Top             =   1275
      Width           =   675
   End
   Begin VB.Label FX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vacío"
      Height          =   240
      Index           =   10
      Left            =   2130
      TabIndex        =   25
      Top             =   1275
      Width           =   675
   End
   Begin VB.Label FX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vacío"
      Height          =   240
      Index           =   11
      Left            =   2910
      TabIndex        =   24
      Top             =   1275
      Width           =   675
   End
   Begin VB.Label FX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vacío"
      Height          =   240
      Index           =   12
      Left            =   555
      TabIndex        =   23
      Top             =   1620
      Width           =   675
   End
   Begin VB.Label FX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vacío"
      Height          =   240
      Index           =   13
      Left            =   1350
      TabIndex        =   22
      Top             =   1620
      Width           =   675
   End
   Begin VB.Label FX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vacío"
      Height          =   240
      Index           =   14
      Left            =   2115
      TabIndex        =   21
      Top             =   1620
      Width           =   675
   End
   Begin VB.Label FX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vacío"
      Height          =   240
      Index           =   15
      Left            =   2925
      TabIndex        =   20
      Top             =   1620
      Width           =   675
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fx Console"
      ForeColor       =   &H00C0C0C0&
      Height          =   225
      Left            =   600
      TabIndex        =   19
      Top             =   270
      Width           =   2940
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Rythm Console"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   600
      TabIndex        =   18
      Top             =   2040
      Width           =   2940
   End
   Begin VB.Label RTM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "detener"
      ForeColor       =   &H0000FF00&
      Height          =   240
      Index           =   0
      Left            =   540
      TabIndex        =   17
      Top             =   2385
      Width           =   675
   End
   Begin VB.Label RTM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vacío"
      Height          =   240
      Index           =   1
      Left            =   1350
      TabIndex        =   16
      Top             =   2385
      Width           =   675
   End
   Begin VB.Label RTM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vacío"
      Height          =   240
      Index           =   2
      Left            =   2130
      TabIndex        =   15
      Top             =   2385
      Width           =   675
   End
   Begin VB.Label RTM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vacío"
      Height          =   240
      Index           =   3
      Left            =   2910
      TabIndex        =   14
      Top             =   2385
      Width           =   675
   End
   Begin VB.Label RTM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vacío"
      Height          =   240
      Index           =   4
      Left            =   540
      TabIndex        =   13
      Top             =   2730
      Width           =   675
   End
   Begin VB.Label RTM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vacío"
      Height          =   240
      Index           =   5
      Left            =   1350
      TabIndex        =   12
      Top             =   2730
      Width           =   675
   End
   Begin VB.Label RTM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vacío"
      Height          =   240
      Index           =   6
      Left            =   2130
      TabIndex        =   11
      Top             =   2730
      Width           =   675
   End
   Begin VB.Label RTM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vacío"
      Height          =   240
      Index           =   7
      Left            =   2910
      TabIndex        =   10
      Top             =   2730
      Width           =   675
   End
   Begin VB.Label RTM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vacío"
      Height          =   240
      Index           =   8
      Left            =   540
      TabIndex        =   9
      Top             =   3075
      Width           =   675
   End
   Begin VB.Label RTM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vacío"
      Height          =   240
      Index           =   9
      Left            =   1350
      TabIndex        =   8
      Top             =   3075
      Width           =   675
   End
   Begin VB.Label RTM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vacío"
      Height          =   240
      Index           =   10
      Left            =   2130
      TabIndex        =   7
      Top             =   3075
      Width           =   675
   End
   Begin VB.Label RTM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vacío"
      Height          =   240
      Index           =   11
      Left            =   2910
      TabIndex        =   6
      Top             =   3075
      Width           =   675
   End
   Begin VB.Label RTM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vacío"
      Height          =   240
      Index           =   12
      Left            =   540
      TabIndex        =   5
      Top             =   3405
      Width           =   675
   End
   Begin VB.Label RTM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vacío"
      Height          =   240
      Index           =   13
      Left            =   1350
      TabIndex        =   4
      Top             =   3405
      Width           =   675
   End
   Begin VB.Label RTM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vacío"
      Height          =   240
      Index           =   14
      Left            =   2130
      TabIndex        =   3
      Top             =   3405
      Width           =   675
   End
   Begin VB.Label RTM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "vacío"
      Height          =   240
      Index           =   15
      Left            =   2910
      TabIndex        =   2
      Top             =   3405
      Width           =   675
   End
   Begin VB.Image image3 
      Enabled         =   0   'False
      Height          =   3975
      Left            =   0
      Picture         =   "FX.frx":0512
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim N As Integer

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 CURX = X
 CURY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 0 Then
  If Label10.ForeColor <> &HFF& Then Label10.ForeColor = &HFF&
  N = 0
  Do While N <= 15
   If FX(N).ForeColor <> UNSELCOLOR Then FX(N).ForeColor = UNSELCOLOR
   If RTM(N).ForeColor <> UNSELCOLOR And RTM(N).ForeColor <> ONCOLOR Then RTM(N).ForeColor = UNSELCOLOR
   N = N + 1
  Loop
 Else
  Me.Move Me.left + (X - CURX), Me.top + (Y - CURY)
 End If
End Sub

Private Sub FX_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 1 Then
  If EFFECT.List(Index) <> "0" Then
   If EFXNAME <> EFFECT.List(Index) Then
    If EFXNAME <> "" Then Call PLAYERCLOSE("EFX")
    Call PLAYEROPEN(hWnd, "EFX", EFFECT.List(Index), "MPEGVideo")
   End If
   If PLAYERSTATUS("EFX") = "playing" Then
    Call PLAYERSETPOS("EFX", 0)
   Else
    Call PLAYERPLAY("EFX", vbNullString, vbNullString)
   End If
  Else
   Form3.Text1 = Trim(str(Index + 1))
   Form3.Text2 = 0
   Form3.Text3 = Index
   If Index > 0 Then Form3.Dir1.path = SOLOPATH(EFFECT.List(Index - 1))
   Form3.Enabled = True
   Form3.Visible = True
   Form1.Enabled = False
  End If
 Else
  Form3.Text2 = 0
  Form3.Text3 = Index
  If FX(Index) = "vacío" Then
   Form3.Text1 = Trim(str(Index + 1))
   If Index > 0 Then Form3.Dir1.path = SOLOPATH(EFFECT.List(Index - 1))
  Else
   Form3.Text1 = Trim(FX(Index))
   Form3.Dir1.path = SOLOPATH(EFFECT.List(Index))
  End If
  N = 0
  Do While N <= Form3.File1.ListCount - 1
   If Form3.File1.List(N) = SOLOFILE(EFFECT.List(Index)) Then Form3.File1.Selected(N) = True
   N = N + 1
  Loop
  Form3.Enabled = True
  Form3.Visible = True
  Form1.Enabled = False
 End If
End Sub

Private Sub FX_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 N = 0
 Do While N <= 15
  If Index = FX(N).Index Then
   If FX(N).ForeColor <> SELCOLOR Then FX(N).ForeColor = SELCOLOR
  Else
   If FX(N).ForeColor <> UNSELCOLOR Then FX(N).ForeColor = UNSELCOLOR
  End If
  N = N + 1
 Loop
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Me.WindowState = 1
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Label10.ForeColor <> SELCOLOR Then Label10.ForeColor = SELCOLOR
End Sub

Private Sub RTM_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 1 Then
  If Index > 0 Then
   If RITMO.List(Index) <> "0" Then
    If RITMNAME <> RITMO.List(Index) Then
     If RITMNAME <> "" Then
      Call PLAYERCLOSE("RITM")
      Call SetAutoRepeat(hWnd, "RITM", vbNullString, vbNullString, True)
     End If
     Call PLAYEROPEN(hWnd, "RITM", RITMO.List(Index), "MPEGVideo")
    End If
    If PLAYERSTATUS("RITM") = "playing" Then
     Call PLAYERSETPOS("RITM", 0)
    Else
     Call PLAYERPLAY("RITM", vbNullString, vbNullString)
    End If
    If RTM(Index).ForeColor <> ONCOLOR Then
     N = 0
     Do While N <= 15
      If RTM(N).ForeColor <> UNSELCOLOR Then RTM(N).ForeColor = UNSELCOLOR
      N = N + 1
     Loop
     RTM(Index).ForeColor = ONCOLOR
    End If
   Else
    Form3.Text2 = 1
    Form3.Text3 = Index
    Form3.Text1 = Trim(str(Index))
    If Index > 0 Then Form3.Dir1.path = SOLOPATH(RITMO.List(Index - 1))
    Form3.Enabled = True
    Form3.Visible = True
    Form1.Enabled = False
   End If
  Else
   If RTM(0).ForeColor <> ONCOLOR Then
    N = 0
    Do While N <= 15
     If RTM(N).ForeColor <> UNSELCOLOR Then RTM(N).ForeColor = UNSELCOLOR
     N = N + 1
    Loop
    RTM(0).ForeColor = ONCOLOR
    Call PLAYERCLOSE("RITM")
    Call SetAutoRepeat(hWnd, "RITM", vbNullString, vbNullString, True)
   End If
  End If
 Else
  Form3.Text2 = 1
  Form3.Text3 = Index
  If RTM(Index) = "vacío" Then
   Form3.Text1 = Trim(str(Index))
   If Index > 0 Then Form3.Dir1.path = SOLOPATH(RITMO.List(Index - 1))
  Else
   Form3.Text1 = Trim(RTM(Index))
   Form3.Dir1.path = SOLOPATH(RITMO.List(Index))
  End If
  N = 0
  Do While N <= Form3.File1.ListCount - 1
   If Form3.File1.List(N) = SOLOFILE(RITMO.List(Index)) Then Form3.File1.Selected(N) = True
   N = N + 1
  Loop
  Form3.Enabled = True
  Form3.Visible = True
  Form1.Enabled = False
 End If
End Sub

Private Sub RTM_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 N = 0
 Do While N <= 15
  If Index = RTM(N).Index Then
   If RTM(N).ForeColor <> SELCOLOR And RTM(N).ForeColor <> ONCOLOR Then RTM(N).ForeColor = SELCOLOR
  Else
   If RTM(N).ForeColor <> UNSELCOLOR And RTM(N).ForeColor <> ONCOLOR Then RTM(N).ForeColor = UNSELCOLOR
  End If
  N = N + 1
 Loop
End Sub
