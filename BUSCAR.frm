VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Dj Box 4 - Buscar"
   ClientHeight    =   3780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   Icon            =   "BUSCAR.frx":0000
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List2 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H0000FF00&
      Height          =   2595
      ItemData        =   "BUSCAR.frx":0442
      Left            =   4155
      List            =   "BUSCAR.frx":0444
      TabIndex        =   3
      Top             =   855
      Visible         =   0   'False
      Width           =   4305
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1650
      TabIndex        =   2
      Text            =   "0"
      Top             =   -75
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   2595
      ItemData        =   "BUSCAR.frx":0446
      Left            =   135
      List            =   "BUSCAR.frx":0448
      TabIndex        =   1
      Top             =   615
      Width           =   4305
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   360
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   4305
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancelar"
      Height          =   210
      Left            =   2430
      TabIndex        =   5
      Top             =   3405
      Width           =   1275
   End
   Begin VB.Label Command2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Aceptar"
      Height          =   210
      Left            =   900
      TabIndex        =   4
      Top             =   3405
      Width           =   1275
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   3780
      Left            =   0
      Picture         =   "BUSCAR.frx":044A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim N As Integer

Private Sub Command1_Click()
 Form1.Enabled = True
 Unload Me
End Sub

Private Sub Command2_Click()
 If Text2 = 0 Then
  If List1.ListIndex > -1 Then
   Form1.LISTA(0).Selected(List2.List(List1.ListIndex)) = True
   Call Form1.CARGAR_A
  Else
   MsgBox ("Debe Escribir El Nombre Del Tema A Buscar")
  End If
 Else
  If List1.ListIndex > -1 Then
   Form1.LISTB(0).Selected(List2.List(List1.ListIndex)) = True
   Call Form1.CARGAR_B
  Else
   MsgBox ("Debe Escribir El Nombre Del Tema A Buscar")
  End If
 End If
 Form1.Enabled = True
 Unload Me
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Command1.ForeColor <> UNSELCOLOR Then Command1.ForeColor = UNSELCOLOR
 If Command2.ForeColor = UNSELCOLOR Then Command2.ForeColor = SELCOLOR
End Sub

Private Sub Form_Load()
  If SKINFILE > 0 Then Call SET_SKIN(6, SKINFILE)
  Text1.ForeColor = DATAFORE
  Text1.BackColor = DATABACK
  List1.ForeColor = DATAFORE
  List1.BackColor = DATABACK
  Command1.ForeColor = UNSELCOLOR
  Command2.ForeColor = UNSELCOLOR
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Command1.ForeColor <> UNSELCOLOR Then Command1.ForeColor = UNSELCOLOR
 If Command2.ForeColor <> UNSELCOLOR Then Command2.ForeColor = UNSELCOLOR
End Sub

Private Sub Text1_Change()
 N = 0
 If Text2 = 0 Then
  List1.Clear
  List2.Clear
  If Text1 <> "" Then
   Do While N <= Form1.LISTA(0).ListCount - 1
    If InStr(LCase(Form1.LISTA(0).List(N)), LCase(Text1)) > 0 Then
     List1.AddItem Form1.LISTA(0).List(N)
     List2.AddItem N
    End If
    N = N + 1
   Loop
  End If
 Else
  List1.Clear
  List2.Clear
  If Text1 <> "" Then
   Do While N <= Form1.LISTB(0).ListCount - 1
    If InStr(LCase(Form1.LISTB(0).List(N)), LCase(Text1)) > 0 Then
     List1.AddItem Form1.LISTB(0).List(N)
     List2.AddItem N
    End If
    N = N + 1
   Loop
  End If
 End If
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Command2.ForeColor <> UNSELCOLOR Then Command2.ForeColor = UNSELCOLOR
 If Command1.ForeColor = UNSELCOLOR Then Command1.ForeColor = SELCOLOR
End Sub
