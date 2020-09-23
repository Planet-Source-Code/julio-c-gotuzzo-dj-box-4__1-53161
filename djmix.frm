VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form10 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form10"
   ClientHeight    =   3570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4860
   LinkTopic       =   "Form10"
   LockControls    =   -1  'True
   ScaleHeight     =   3570
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox OPT 
      Caption         =   "Check1"
      Height          =   225
      Index           =   3
      Left            =   2520
      TabIndex        =   10
      Top             =   3120
      Width           =   210
   End
   Begin VB.CheckBox OPT 
      Caption         =   "Check1"
      Height          =   225
      Index           =   2
      Left            =   2520
      TabIndex        =   6
      Top             =   2760
      Width           =   210
   End
   Begin VB.CheckBox OPT 
      Caption         =   "Check1"
      Height          =   225
      Index           =   1
      Left            =   480
      TabIndex        =   5
      Top             =   3120
      Width           =   210
   End
   Begin VB.CheckBox OPT 
      Caption         =   "Check1"
      Height          =   225
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Top             =   2760
      Width           =   210
   End
   Begin ComctlLib.Slider MIXVEL 
      Height          =   255
      Left            =   1035
      TabIndex        =   0
      Top             =   1470
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   450
      _Version        =   327682
      SelStart        =   7
      Value           =   7
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Break Fade"
      Height          =   240
      Index           =   7
      Left            =   525
      TabIndex        =   23
      Top             =   1020
      Width           =   900
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Break Mix"
      Height          =   240
      Index           =   6
      Left            =   1530
      TabIndex        =   22
      Top             =   1020
      Width           =   900
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pitch Sch."
      Height          =   240
      Index           =   5
      Left            =   2505
      TabIndex        =   21
      Top             =   1020
      Width           =   900
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fx Intro"
      Height          =   240
      Index           =   4
      Left            =   3480
      TabIndex        =   20
      Top             =   1020
      Width           =   900
   End
   Begin VB.Label Label7 
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
      Left            =   4545
      TabIndex        =   19
      ToolTipText     =   "Minimizar"
      Top             =   45
      Width           =   255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Play Fade In"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   3135
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Stop Fade Out"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   2775
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Dj Efx"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   2400
      Width           =   3375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Dj Mix"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Vel."
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   660
      TabIndex        =   1
      Top             =   1470
      Width           =   360
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Fx Fade Out"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   2880
      TabIndex        =   9
      Top             =   2775
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Deck End Fade"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   2865
      TabIndex        =   18
      Top             =   3135
      Width           =   1335
   End
   Begin VB.Label RANBUT 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Auto Dj"
      Height          =   255
      Left            =   1980
      TabIndex        =   17
      Top             =   1950
      Width           =   945
   End
   Begin VB.Label PANBUT 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "P"
      Height          =   240
      Left            =   3930
      TabIndex        =   16
      ToolTipText     =   "0"
      Top             =   1500
      Width           =   315
   End
   Begin VB.Label LOOPBUT 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      Height          =   240
      Left            =   3495
      TabIndex        =   15
      ToolTipText     =   "0"
      Top             =   1500
      Width           =   315
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CrossFade"
      Height          =   240
      Index           =   3
      Left            =   3480
      TabIndex        =   14
      Top             =   660
      Width           =   900
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fade Sch."
      Height          =   240
      Index           =   2
      Left            =   2505
      TabIndex        =   13
      Top             =   660
      Width           =   900
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fade Mix"
      Height          =   240
      Index           =   1
      Left            =   1530
      TabIndex        =   12
      Top             =   660
      Width           =   900
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Switch"
      Height          =   240
      Index           =   0
      Left            =   525
      TabIndex        =   11
      Top             =   660
      Width           =   900
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   3570
      Left            =   0
      Picture         =   "djmix.frx":0000
      Top             =   0
      Width           =   4860
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 LOOPBUT.ToolTipText = 0
 PANBUT.ToolTipText = 1
 Call SETMIXING(Index)
End Sub

Private Sub Command1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 n = 0
 Do While n <= 3
  If Index = Command1(n).Index Then
   If Command1(n).ForeColor <> SELCOLOR And Command1(n).ForeColor <> ONCOLOR Then Command1(n).ForeColor = SELCOLOR
  Else
   If Command1(n).ForeColor <> UNSELCOLOR And Command1(n).ForeColor <> ONCOLOR Then Command1(n).ForeColor = UNSELCOLOR
  End If
  n = n + 1
 Loop
End Sub

Private Sub Form_Load()
 Randomize
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 CURX = X
 CURY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 0 Then
  If Label7.ForeColor = &HFFFF& Then Label7.ForeColor = &HFF&
  If RANBUT.ForeColor <> UNSELCOLOR And RANBUT.ForeColor <> ONCOLOR Then RANBUT.ForeColor = UNSELCOLOR
  If LOOPBUT.ForeColor <> UNSELCOLOR And LOOPBUT.ForeColor <> ONCOLOR Then LOOPBUT.ForeColor = UNSELCOLOR
  If PANBUT.ForeColor <> UNSELCOLOR And PANBUT.ForeColor <> ONCOLOR Then PANBUT.ForeColor = UNSELCOLOR
  n = 0
  Do While n <= 3
   If Command1(n).ForeColor <> UNSELCOLOR And Command1(n).ForeColor <> ONCOLOR Then Command1(n).ForeColor = UNSELCOLOR
   n = n + 1
  Loop
 Else
  Me.Move Me.left + (X - CURX), Me.top + (Y - CURY)
 End If
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Me.WindowState = 1
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Label7.ForeColor <> &HFFFF& Then Label7.ForeColor = &HFFFF&
End Sub

Private Sub LOOPBUT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If LOOPBUT.ForeColor = ONCOLOR Then
  LOOPBUT.ForeColor = SELCOLOR
 Else
  LOOPBUT.ForeColor = ONCOLOR
 End If
End Sub

Private Sub LOOPBUT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If LOOPBUT.ForeColor <> ONCOLOR And LOOPBUT.ForeColor <> SELCOLOR Then LOOPBUT.ForeColor = SELCOLOR
End Sub

Private Sub PANBUT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If PANBUT.ForeColor = ONCOLOR Then
  PANBUT.ForeColor = SELCOLOR
 Else
  PANBUT.ForeColor = ONCOLOR
 End If
End Sub

Private Sub PANBUT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If PANBUT.ForeColor <> ONCOLOR And PANBUT.ForeColor <> SELCOLOR Then PANBUT.ForeColor = SELCOLOR
End Sub

Private Sub RANBUT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' RANBUT.ForeColor = SELCOLOR
' PASS = InputBox("Inserte un password para bloquear el Automix:")
' If Trim(PASS) = "" Then Exit Sub
' Form1.Enabled = False
' Form10.Enabled = False
' Form4.Enabled = False
' Form5.Enabled = False
' Form8.Enabled = False
' Do While InputBox("Inserte el password para desbloquear el Automix:") <> PASS
'  DoEvents
' Loop
' Form1.Enabled = True
' Form10.Enabled = True
' Form4.Enabled = True
' Form5.Enabled = True
' Form8.Enabled = True
' RANBUT.ForeColor = UNSELCOLOR
 Form9.Enabled = True
 Form9.Visible = True
End Sub

Private Sub RANBUT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If RANBUT.ForeColor <> ONCOLOR And RANBUT.ForeColor <> SELCOLOR Then RANBUT.ForeColor = SELCOLOR
End Sub

Private Sub SETMIXING(MixNumber As Integer)
 Select Case MixNumber
 Case Is = 0
  Call Form1.SWITCH
 Case Is = 1
  If MIXSET(1) = 0 Then
   If PLAYERSTATUS("DECK1") <> "playing" And PLAYERSTATUS("DECK2") <> "stopped" Then
    MIXSET(1) = 12
   Else
    If PLAYERSTATUS("DECK1") <> "stopped" And PLAYERSTATUS("DECK2") <> "playing" Then
     MIXSET(1) = 11
    End If
   End If
  End If
 Case Is = 2
  If MIXSET(1) = 0 Then
   If PLAYERSTATUS("DECK1") <> "playing" And PLAYERSTATUS("DECK2") <> "stopped" Then
    MIXSET(1) = 22
   Else
    If PLAYERSTATUS("DECK1") <> "stopped" And PLAYERSTATUS("DECK2") <> "playing" Then
     MIXSET(1) = 21
    End If
   End If
  End If
 Case Is = 3
  If MIXSET(1) = 0 Then
   If PLAYERSTATUS("DECK1") <> "playing" And PLAYERSTATUS("DECK2") <> "stopped" Then
    MIXSET(1) = 32
   Else
    If PLAYERSTATUS("DECK1") <> "stopped" And PLAYERSTATUS("DECK2") <> "playing" Then
     MIXSET(1) = 31
    End If
   End If
  End If
 End Select
End Sub
