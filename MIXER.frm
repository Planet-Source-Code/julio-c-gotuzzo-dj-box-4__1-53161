VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form5 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Dj Box 4 - Mixer"
   ClientHeight    =   3660
   ClientLeft      =   495
   ClientTop       =   4995
   ClientWidth     =   6450
   Icon            =   "MIXER.frx":0000
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.Slider CROSSFADE 
      Height          =   360
      Left            =   2670
      TabIndex        =   0
      Top             =   3075
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   635
      _Version        =   327682
      Max             =   200
      SelectRange     =   -1  'True
      TickStyle       =   3
      Value           =   100
   End
   Begin ComctlLib.Slider VOL 
      Height          =   2580
      Index           =   0
      Left            =   1245
      TabIndex        =   1
      Top             =   465
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   4551
      _Version        =   327682
      Orientation     =   1
      Max             =   100
      SelectRange     =   -1  'True
      TickStyle       =   3
   End
   Begin ComctlLib.Slider VOL 
      Height          =   1950
      Index           =   1
      Left            =   2700
      TabIndex        =   2
      Top             =   465
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   3440
      _Version        =   327682
      Orientation     =   1
      Max             =   100
      SelectRange     =   -1  'True
      TickStyle       =   3
   End
   Begin ComctlLib.Slider VOL 
      Height          =   1950
      Index           =   2
      Left            =   3345
      TabIndex        =   3
      Top             =   465
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   3440
      _Version        =   327682
      Orientation     =   1
      Max             =   100
      SelectRange     =   -1  'True
      TickStyle       =   3
   End
   Begin ComctlLib.Slider VOL 
      Height          =   1950
      Index           =   3
      Left            =   3990
      TabIndex        =   4
      Top             =   465
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   3440
      _Version        =   327682
      Orientation     =   1
      Max             =   100
      SelectRange     =   -1  'True
      TickStyle       =   3
   End
   Begin ComctlLib.Slider VOL 
      Height          =   1950
      Index           =   4
      Left            =   4620
      TabIndex        =   5
      Top             =   465
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   3440
      _Version        =   327682
      Orientation     =   1
      Max             =   100
      SelectRange     =   -1  'True
      TickStyle       =   3
   End
   Begin ComctlLib.Slider VOL 
      Height          =   1950
      Index           =   5
      Left            =   5265
      TabIndex        =   18
      Top             =   465
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   3440
      _Version        =   327682
      Orientation     =   1
      Max             =   100
      SelectRange     =   -1  'True
      TickStyle       =   3
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "B"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   22
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "A"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2355
      TabIndex        =   21
      Top             =   3120
      Width           =   255
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   5520
      X2              =   5520
      Y1              =   3120
      Y2              =   3480
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   2835
      X2              =   2835
      Y1              =   3105
      Y2              =   3480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      X1              =   4170
      X2              =   4170
      Y1              =   3090
      Y2              =   3480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cd"
      Enabled         =   0   'False
      Height          =   225
      Left            =   5235
      TabIndex        =   20
      Top             =   195
      Width           =   405
   End
   Begin VB.Label MUTE 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   150
      Index           =   5
      Left            =   5340
      TabIndex        =   19
      ToolTipText     =   "Silenciar."
      Top             =   2595
      Width           =   225
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
      Left            =   6150
      TabIndex        =   17
      ToolTipText     =   "Minimizar"
      Top             =   45
      Width           =   255
   End
   Begin VB.Label MUTE 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   150
      Index           =   0
      Left            =   1335
      TabIndex        =   16
      ToolTipText     =   "Silenciar."
      Top             =   3270
      Width           =   225
   End
   Begin VB.Label MUTE 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   150
      Index           =   1
      Left            =   2775
      TabIndex        =   15
      ToolTipText     =   "Silenciar."
      Top             =   2595
      Width           =   225
   End
   Begin VB.Label MUTE 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   150
      Index           =   2
      Left            =   3420
      TabIndex        =   14
      ToolTipText     =   "Silenciar."
      Top             =   2595
      Width           =   225
   End
   Begin VB.Label MUTE 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   150
      Index           =   3
      Left            =   4080
      TabIndex        =   13
      ToolTipText     =   "Silenciar."
      Top             =   2595
      Width           =   225
   End
   Begin VB.Label MUTE 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   150
      Index           =   4
      Left            =   4695
      TabIndex        =   12
      ToolTipText     =   "Silenciar."
      Top             =   2595
      Width           =   225
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Master"
      Height          =   225
      Left            =   1020
      TabIndex        =   11
      Top             =   180
      Width           =   840
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Deck A"
      Height          =   225
      Left            =   2625
      TabIndex        =   10
      Top             =   195
      Width           =   570
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Deck B"
      Height          =   225
      Left            =   3270
      TabIndex        =   9
      Top             =   195
      Width           =   570
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fx"
      Height          =   225
      Left            =   3885
      TabIndex        =   8
      Top             =   195
      Width           =   570
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Rythm"
      Height          =   225
      Left            =   4515
      TabIndex        =   7
      Top             =   195
      Width           =   570
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cross-Fader"
      Height          =   225
      Left            =   3630
      TabIndex        =   6
      Top             =   2850
      Width           =   1095
   End
   Begin VB.Image Image2 
      Enabled         =   0   'False
      Height          =   3660
      Left            =   0
      Picture         =   "MIXER.frx":0442
      Top             =   0
      Width           =   6450
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim N As Integer

Private Sub CROSSFADE_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If MIXSET(5) > 0 Then MIXSET(5) = 0
End Sub

Private Sub CROSSFADE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Label8.ForeColor <> SELCOLOR Then Label8.ForeColor = SELCOLOR
End Sub

Public Sub CROSSFADE_Scroll()
 If CROSSFADE.Value >= 100 And Label9.ForeColor <> &HFFFF& Then Label9.ForeColor = &HFFFF&
 If CROSSFADE.Value >= 50 And CROSSFADE.Value < 100 And Label9.ForeColor <> &HC0& Then Label9.ForeColor = &HC0&
 If CROSSFADE.Value >= 0 And CROSSFADE.Value < 50 And Label9.ForeColor <> 0 Then Label9.ForeColor = 0
 If CROSSFADE.Value <= 100 And Label2.ForeColor <> &HFFFF& Then Label2.ForeColor = &HFFFF&
 If CROSSFADE.Value <= 150 And CROSSFADE.Value > 100 And Label2.ForeColor <> &HC0& Then Label2.ForeColor = &HC0&
 If CROSSFADE.Value <= 200 And CROSSFADE.Value > 150 And Label2.ForeColor <> 0 Then Label2.ForeColor = 0
 If CROSSFADE.Value <> 100 Then Line1.BorderColor = &H0& Else Line1.BorderColor = &HFFFF&
 If CROSSFADE.Value > 0 Then Line2.BorderColor = &H0& Else Line2.BorderColor = &HFFFF&
 If CROSSFADE.Value < 200 Then Line3.BorderColor = &H0& Else Line3.BorderColor = &HFFFF&
 Call VOL1
 Call VOL2
 Form1.ACTION(0).Caption = Trim(str(PLAYERGETVOLUME("DECK1", "all"))) + "%"
 Form1.ACTION(1).Caption = Trim(str(PLAYERGETVOLUME("DECK2", "all"))) + "%"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 CURX = X
 CURY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 0 Then
  If Label10.ForeColor = &HFFFF& Then Label10.ForeColor = &HFF&
  If Label3.ForeColor = SELCOLOR Then Label3.ForeColor = 0
  If Label4.ForeColor = SELCOLOR Then Label4.ForeColor = 0
  If Label5.ForeColor = SELCOLOR Then Label5.ForeColor = 0
  If Label6.ForeColor = SELCOLOR Then Label6.ForeColor = 0
  If Label7.ForeColor = SELCOLOR Then Label7.ForeColor = 0
  If Label8.ForeColor = SELCOLOR Then Label8.ForeColor = 0
 Else
  Me.Move Me.left + (X - CURX), Me.top + (Y - CURY)
 End If
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Me.WindowState = 1
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Label10.ForeColor <> &HFFFF& Then Label10.ForeColor = &HFFFF&
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 MIXSET(5) = 1
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 MIXSET(5) = 2
End Sub

Private Sub MUTE_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 Select Case Index
 Case Is = 0
  If MUTE(Index).BackColor <> &HFF& Then
   MUTE(Index).BackColor = &HFF&
   If MUTE(1).BackColor = &HFF00& Then Call PLAYERMUTE("DECK1", "all", "off")
   If MUTE(2).BackColor = &HFF00& Then Call PLAYERMUTE("DECK2", "all", "off")
   If MUTE(3).BackColor = &HFF00& Then Call PLAYERMUTE("EFX", "all", "off")
   If MUTE(4).BackColor = &HFF00& Then Call PLAYERMUTE("RITM", "all", "off")
  Else
   MUTE(Index).BackColor = &HFF00&
   If MUTE(1).BackColor = &HFF00& Then Call PLAYERMUTE("DECK1", "all", "on")
   If MUTE(2).BackColor = &HFF00& Then Call PLAYERMUTE("DECK2", "all", "on")
   If MUTE(3).BackColor = &HFF00& Then Call PLAYERMUTE("EFX", "all", "on")
   If MUTE(4).BackColor = &HFF00& Then Call PLAYERMUTE("RITM", "all", "on")
  End If
 Case Is = 1
  If MUTE(Index).BackColor <> &HFF& Then
   MUTE(Index).BackColor = &HFF&
   Call PLAYERMUTE("DECK1", "all", "off")
  Else
   MUTE(Index).BackColor = &HFF00&
   If MUTE(0).BackColor <> &HFF& Then Call PLAYERMUTE("DECK1", "all", "on")
  End If
 Case Is = 2
  If MUTE(Index).BackColor <> &HFF& Then
   MUTE(Index).BackColor = &HFF&
   Call PLAYERMUTE("DECK2", "all", "off")
  Else
   MUTE(Index).BackColor = &HFF00&
   If MUTE(0).BackColor <> &HFF& Then Call PLAYERMUTE("DECK2", "all", "on")
  End If
 Case Is = 3
  If MUTE(Index).BackColor <> &HFF& Then
   MUTE(Index).BackColor = &HFF&
   Call PLAYERMUTE("EFX", "all", "off")
  Else
   MUTE(Index).BackColor = &HFF00&
   If MUTE(0).BackColor <> &HFF& Then Call PLAYERMUTE("EFX", "all", "on")
  End If
 Case Is = 4
  If MUTE(Index).BackColor <> &HFF& Then
   MUTE(Index).BackColor = &HFF&
   Call PLAYERMUTE("RITM", "all", "off")
  Else
   MUTE(Index).BackColor = &HFF00&
   If MUTE(0).BackColor <> &HFF& Then Call PLAYERMUTE("RITM", "all", "on")
  End If
 End Select
End Sub

Private Sub VOL_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 VOL(Index).ToolTipText = 100 - VOL(Index).Value
 Select Case Index
 Case Is = 0
  If Label3.ForeColor <> SELCOLOR Then Label3.ForeColor = SELCOLOR
 Case Is = 1
  If Label4.ForeColor <> SELCOLOR Then Label4.ForeColor = SELCOLOR
 Case Is = 2
  If Label5.ForeColor <> SELCOLOR Then Label5.ForeColor = SELCOLOR
 Case Is = 3
  If Label6.ForeColor <> SELCOLOR Then Label6.ForeColor = SELCOLOR
 Case Is = 4
  If Label7.ForeColor <> SELCOLOR Then Label7.ForeColor = SELCOLOR
 End Select
End Sub

Private Sub VOL_Scroll(Index As Integer)
 Select Case Index
 Case Is = 0
  Call VOL1
  Call VOL2
  Call VOL3
  Call VOL4
  Form1.ACTION(0).Caption = Trim(str(PLAYERGETVOLUME("DECK1", "all"))) + "%"
  Form1.ACTION(1).Caption = Trim(str(PLAYERGETVOLUME("DECK2", "all"))) + "%"
 Case Is = 1
  Call VOL1
  Form1.ACTION(0).Caption = Trim(str(PLAYERGETVOLUME("DECK1", "all"))) + "%"
 Case Is = 2
  Call VOL2
  Form1.ACTION(1).Caption = Trim(str(PLAYERGETVOLUME("DECK2", "all"))) + "%"
 Case Is = 3
  Call VOL3
 Case Is = 4
  Call VOL4
 End Select
End Sub

Function VOL1()
 If CROSSFADE.Value <= 100 Then
  Call PLAYERSETVOLUME("DECK1", "all", Int(((100 - VOL(0).Value) * (100 - VOL(1).Value)) / 100))
 Else
  Call PLAYERSETVOLUME("DECK1", "all", Int(((100 - (CROSSFADE.Value - 100)) * Int(((100 - VOL(0).Value) * (100 - VOL(1).Value)) / 100)) / 100))
 End If
End Function

Function VOL2()
 If CROSSFADE.Value >= 100 Then
  Call PLAYERSETVOLUME("DECK2", "all", Int(((100 - VOL(0).Value) * (100 - VOL(2).Value)) / 100))
 Else
  Call PLAYERSETVOLUME("DECK2", "all", Int((CROSSFADE.Value * Int(((100 - VOL(0).Value) * (100 - VOL(2).Value)) / 100)) / 100))
 End If
End Function

Function VOL3()
 Call PLAYERSETVOLUME("EFX", "all", Int(((100 - VOL(0).Value) * (100 - VOL(3).Value)) / 100))
End Function

Function VOL4()
 Call PLAYERSETVOLUME("RITM", "all", Int(((100 - VOL(0).Value) * (100 - VOL(4).Value)) / 100))
End Function
