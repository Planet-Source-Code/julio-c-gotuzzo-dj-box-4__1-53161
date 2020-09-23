VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form Form8 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Dj Box 4 - Cd Player"
   ClientHeight    =   2385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5595
   LinkTopic       =   "Form8"
   LockControls    =   -1  'True
   ScaleHeight     =   2385
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin PicClip.PictureClip DISCCLIP 
      Left            =   4725
      Top             =   1890
      _ExtentX        =   3651
      _ExtentY        =   2170
      _Version        =   393216
      Rows            =   2
      Cols            =   3
      Picture         =   "CdPlayer.frx":0000
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H0000FFFF&
      Height          =   150
      Left            =   135
      ScaleHeight     =   90
      ScaleWidth      =   105
      TabIndex        =   13
      ToolTipText     =   "Cerrar / Abrir Dispositivo."
      Top             =   90
      Width           =   165
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   1800
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      ItemData        =   "CdPlayer.frx":17A2
      Left            =   990
      List            =   "CdPlayer.frx":17A4
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1890
      Width           =   3645
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   210
      Left            =   915
      TabIndex        =   1
      Top             =   1215
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   370
      _Version        =   327682
      TickStyle       =   3
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00000007&
      Height          =   915
      Left            =   825
      ScaleHeight     =   855
      ScaleWidth      =   3900
      TabIndex        =   0
      Top             =   240
      Width           =   3960
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CDBASE"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   165
         Left            =   0
         TabIndex        =   14
         Top             =   675
         Width           =   765
      End
      Begin VB.Image STATEBMP 
         Height          =   150
         Left            =   30
         Stretch         =   -1  'True
         Top             =   30
         Width           =   165
      End
      Begin VB.Image DISCBMP 
         Height          =   270
         Left            =   3540
         Stretch         =   -1  'True
         Top             =   540
         Width           =   315
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "- | 00:00:00 |"
         ForeColor       =   &H0000FF00&
         Height          =   225
         Left            =   2280
         TabIndex        =   12
         Top             =   480
         Width           =   1005
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "< 00:00:00 >"
         ForeColor       =   &H0000FF00&
         Height          =   225
         Left            =   1320
         TabIndex        =   11
         Top             =   480
         Width           =   1005
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "No Se Encuentra Lectora Disponible..."
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   1095
         TabIndex        =   10
         Top             =   120
         Width           =   2745
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   195
         Width           =   855
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Eject"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   4
      Left            =   4095
      TabIndex        =   8
      ToolTipText     =   "Abrir/Cerrar Bandeja."
      Top             =   1575
      Width           =   780
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Stop"
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   3
      Left            =   3240
      TabIndex        =   7
      ToolTipText     =   "Detener."
      Top             =   1575
      Width           =   810
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Next"
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   2
      Left            =   2400
      TabIndex        =   6
      ToolTipText     =   "Pista Siguiente."
      Top             =   1575
      Width           =   810
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Prev."
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   1
      Left            =   1590
      TabIndex        =   5
      ToolTipText     =   "Pista Anterior."
      Top             =   1575
      Width           =   780
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Play"
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   0
      Left            =   765
      TabIndex        =   4
      ToolTipText     =   "Reproducir/Pausar Pista Actual."
      Top             =   1575
      Width           =   765
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
      Height          =   225
      Left            =   5325
      TabIndex        =   3
      ToolTipText     =   "Minimizar"
      Top             =   15
      Width           =   225
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   2445
      Left            =   -75
      Picture         =   "CdPlayer.frx":17A6
      Top             =   -15
      Width           =   5715
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tempor As String
Dim LOCK3 As Boolean
Dim n As Integer

Private Sub Combo1_Click()
 Dim n As Integer
 If MediaPresent = True Then
  n = 0
  If IsPlaying = True Then
   n = 1
   StopCD
  End If
  If SetCurrentTrack(Combo1.ListIndex + 1) = False Then MsgBox ("Error al intentar acceder a pista " + PISTA(Val(Combo1.ListIndex + 1), 4) + ".")
  Call INIT_CD
  If n = 1 Then
   PlayCD
   Label1(0) = "Pause"
  End If
 End If
End Sub

Private Sub Form_Load()
 DISCBMP.Picture = DISCCLIP.GraphicCell(5)
 STATEBMP.Picture = Form1.STATECLIP.GraphicCell(1)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 CURX = X
 CURY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 0 Then
  If Label10.ForeColor = &HFFFF& Then Label10.ForeColor = &HFF&
  n = 0
  Do While n <= 4
   If Label1(n).ForeColor <> UNSELCOLOR And Label1(n).ForeColor <> ONCOLOR Then Label1(n).ForeColor = UNSELCOLOR
   n = n + 1
  Loop
 Else
  Me.Move Me.left + (X - CURX), Me.top + (Y - CURY)
 End If
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim n As Integer
 Dim capo As Long
 Select Case Index
 Case Is = 0
  If MediaPresent = True Then
   If Label1(0).Caption = "Play" Then
    If PlayCD = False Then MsgBox ("Error al intentar reproducir la pista " + PISTA(Val(left(GetCurrentPosition, 2)), 4) + ".") Else Label1(0) = "Pause"
   Else
    If Label1(0).ForeColor = ONCOLOR Then
     If PlayCD = False Then MsgBox ("Error al intentar reproducir la pista " + PISTA(Val(left(GetCurrentPosition, 2)), 4) + ".") Else Label1(0).ForeColor = SELCOLOR
    Else
     If PauseCD = False Then MsgBox ("Error al intentar pausar la pista " + PISTA(Val(left(GetCurrentPosition, 2)), 4) + ".") Else Label1(0).ForeColor = ONCOLOR
     If STATEBMP.Picture <> Form1.STATECLIP.GraphicCell(2) Then STATEBMP.Picture = Form1.STATECLIP.GraphicCell(2)
    End If
   End If
  End If
 Case Is = 1
  If MediaPresent = True And Val(left(GetCurrentPosition, 2)) > 1 Then
   Label1(0).ForeColor = UNSELCOLOR
   Label1(0).Caption = "Play"
   n = 0
   capo = Val(left(GetCurrentPosition, 2)) - 1
   If IsPlaying = True Then
    StopCD
    n = 1
   End If
   If SetCurrentTrack(capo) = False Then MsgBox ("Error al intentar acceder a pista " + PISTA(Val(left(GetCurrentPosition, 2)) - 1, 4) + ".")
   Call INIT_CD
   If n = 1 Then
    PlayCD
    Label1(0) = "Pause"
   End If
  End If
 Case Is = 2
  If MediaPresent = True And Val(left(GetCurrentPosition, 2)) < GetNumberOfTracks Then
   Label1(0).Caption = "Play"
   Label1(0).ForeColor = UNSELCOLOR
   n = 0
   capo = Val(left(GetCurrentPosition, 2)) + 1
   If IsPlaying = True Then
    StopCD
    n = 1
   End If
   If SetCurrentTrack(capo) = False Then MsgBox ("Error al intentar acceder a pista " + PISTA(Val(left(GetCurrentPosition, 2)) + 1, 4) + ".")
   Call INIT_CD
   If n = 1 Then
    PlayCD
    Label1(0) = "Pause"
   End If
  End If
 Case Is = 3
  If MediaPresent = True Then
   If StopCD = False Then
    MsgBox ("Error al intentar detener la pista " + PISTA(Val(left(GetCurrentPosition, 2)), 4) + ".")
   Else
    Label1(0).Caption = "Play"
    If STATEBMP.Picture = Form1.STATECLIP.GraphicCell(2) Then STATEBMP.Picture = Form1.STATECLIP.GraphicCell(1)
    Label1(0).ForeColor = UNSELCOLOR
    SetCurrentTrack (Val(Trim(Label2)))
    Call INIT_CD
   End If
  End If
 Case Is = 4
  If mciOpen = True Then
   Label1(0).Caption = "Play"
   Label1(0).ForeColor = UNSELCOLOR
   If MediaPresent = True And drawer = True Then drawer = False
   If drawer = True Then
    Label3 = "Inicializando..."
    ShutCD
    SetCurrentTrack (1)
   Else
    Label3 = "Inserte Un CD De Audio."
    EjectCD
   End If
  End If
 End Select
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 n = 0
 Do While n <= 4
  If Index = Label1(n).Index Then
   If Label1(n).ForeColor <> SELCOLOR And Label1(n).ForeColor <> ONCOLOR Then Label1(n).ForeColor = SELCOLOR
  Else
   If Label1(n).ForeColor <> UNSELCOLOR And Label1(n).ForeColor <> ONCOLOR Then Label1(n).ForeColor = UNSELCOLOR
  End If
  n = n + 1
 Loop
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Me.WindowState = 1
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Label10.ForeColor <> &HFFFF& Then Label10.ForeColor = &HFFFF&
End Sub

Private Sub Label6_Click()
 Form12.Enabled = True
 Form12.Visible = True
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Picture2.BackColor = &HFFFF& Then
  Picture2.BackColor = &HFF&
  Timer1.Enabled = False
  If IsStopped = False Then StopCD
  CloseCD
  Label3 = "Dispositivo Cerrado."
  If STATEBMP.Picture <> Form1.STATECLIP.GraphicCell(1) Then STATEBMP.Picture = Form1.STATECLIP.GraphicCell(1)
  If Label2 <> "0000" Then Label2 = "0000"
  If Label4 <> "< 00:00:00 >" Then Label4 = "< 00:00:00 >"
  If Label5 <> "- | 00:00:00 |" Then Label5 = "- | 00:00:00 |"
 Else
  If OpenCD("d:\") = False Then
   If OpenCD("e:\") = False Then
    If OpenCD("f:\") = False Then
     If OpenCD("g:\") = False Then
      If OpenCD("h:\") = False Then MsgBox ("No Hay Una Lectora Presente En El Sistema o Está Siendo Utilizada Por Otra Aplicación.")
     End If
    End If
   End If
  End If
  If mciOpen = True Then
   Picture2.BackColor = &HFFFF&
   Timer1.Enabled = True
   If MediaPresent = True Then INIT_CD
  End If
 End If
End Sub

Private Sub Slider1_Change()
 Dim capo As Long
 Dim n As Integer
 If LOCK3 = True Then
  n = 0
  capo = Val(left(GetCurrentPosition, 2))
  If IsPlaying = True Then
   StopCD
   n = 1
  End If
  SetCurrentTime (Trim(str(capo)) + ":" + MINSEG(Slider1.Value) + ":00")
  If n = 1 Then PlayCD
  LOCK3 = False
 End If
End Sub

Private Sub Slider1_Scroll()
 If LOCK3 = False Then LOCK3 = True
 Label4.Caption = MINSEG(Slider1.Value) + ":00"
End Sub

Private Sub Timer1_Timer()
  If mciOpen = True Then
   If MediaPresent = True Then
    If IsPlaying = True Then
     If STATEBMP.Picture <> Form1.STATECLIP.GraphicCell(0) Then STATEBMP.Picture = Form1.STATECLIP.GraphicCell(0)
    Else
     If STATEBMP.Picture <> Form1.STATECLIP.GraphicCell(1) And STATEBMP.Picture <> Form1.STATECLIP.GraphicCell(2) Then STATEBMP.Picture = Form1.STATECLIP.GraphicCell(1)
    End If
    If LOCK3 = False Then
     tempor = GetCurrentPosition
     
     If Trim(str(left(tempor, 2))) <> Trim(str(Label2)) Then Call INIT_CD
     If Label4 <> "< " + right(tempor, 8) + " >" Then Label4 = "< " + right(tempor, 8) + " >"
     If Slider1.Value <> (Val(Mid(right(tempor, 8), 1, 2)) * 60) + Val(Mid(right(tempor, 8), 4, 2)) Then Slider1.Value = (Val(Mid(right(tempor, 8), 1, 2)) * 60) + Val(Mid(right(tempor, 8), 4, 2))
    End If
    If Combo1.ListCount <= 0 Then INIT_CD
   Else
    If Combo1.ListCount > 0 Then Combo1.Clear
    If STATEBMP.Picture <> Form1.STATECLIP.GraphicCell(1) Then STATEBMP.Picture = Form1.STATECLIP.GraphicCell(1)
    If DISCBMP.Picture <> DISCCLIP.GraphicCell(5) Then DISCBMP.Picture = DISCCLIP.GraphicCell(5)
    Label3 = "Inserte Un CD De Audio."
    If Label2 <> "0000" Then Label2 = "0000"
    If Label4 <> "< 00:00:00 >" Then Label4 = "< 00:00:00 >"
    If Label5 <> "- | 00:00:00 |" Then Label5 = "- | 00:00:00 |"
   End If
  Else
   If Combo1.ListCount > 0 Then Combo1.Clear
   Label3 = "No Se Encuentra Dispositivo."
   If STATEBMP.Picture <> Form1.STATECLIP.GraphicCell(1) Then STATEBMP.Picture = Form1.STATECLIP.GraphicCell(1)
   If DISCBMP.Picture <> DISCCLIP.GraphicCell(5) Then DISCBMP.Picture = DISCCLIP.GraphicCell(5)
   If Label2 <> "0000" Then Label2 = "0000"
   If Label4 <> "< 00:00:00 >" Then Label4 = "< 00:00:00 >"
   If Label5 <> "- | 00:00:00 |" Then Label5 = "- | 00:00:00 |"
  End If
End Sub

Public Sub INIT_CD()
 Label1(0).Caption = "Play"
 Label1(0).ForeColor = UNSELCOLOR
 If Label2 <> PISTA(Val(left(GetCurrentPosition, 2)), 4) Then Label2 = PISTA(Val(left(GetCurrentPosition, 2)), 4)
 If Label5 <> "- | " + GetTrackLength(Val(left(GetCurrentPosition, 2))) + " |" Then Label5 = "- | " + GetTrackLength(Val(left(GetCurrentPosition, 2))) + " |"
 Slider1.Max = (Val(Mid(GetTrackLength(Val(left(GetCurrentPosition, 2))), 1, 2)) * 60) + Val(Mid(GetTrackLength(Val(left(GetCurrentPosition, 2))), 4, 2))
 Label3 = Chr(34) + " " + Combo1.List(Combo1.ListIndex) + " " + Chr(34)
 If Combo1.ListCount <> GetNumberOfTracks Then
  Combo1.Clear
 If LoadCDBase(Combo1, GetCDID, Combo1, Combo1, 1) = False Then
   Combo1.Clear
   Do While Combo1.ListCount < GetNumberOfTracks
    Combo1.AddItem "Pista " + str(Combo1.ListCount + 1)
   Loop
 End If
 End If
 If Combo1.ListIndex + 1 <> Val(left(GetCurrentPosition, 2)) Then Combo1.ListIndex = Val(left(GetCurrentPosition, 2)) - 1
 If DISCBMP.Picture <> DISCCLIP.GraphicCell(0) Then DISCBMP.Picture = DISCCLIP.GraphicCell(0)
End Sub
