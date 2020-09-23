VERSION 5.00
Begin VB.Form Form9 
   BorderStyle     =   0  'None
   Caption         =   "Dj Box 4 - Lights"
   ClientHeight    =   3495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4455
   LinkTopic       =   "Form9"
   LockControls    =   -1  'True
   ScaleHeight     =   3495
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox ENC 
      BackColor       =   &H000000FF&
      Height          =   150
      Left            =   120
      ScaleHeight     =   90
      ScaleWidth      =   105
      TabIndex        =   7
      ToolTipText     =   "Encender/Apagar Dispositivo"
      Top             =   105
      Width           =   165
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   90
      Top             =   2775
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   1620
      ItemData        =   "Light.frx":0000
      Left            =   495
      List            =   "Light.frx":0002
      TabIndex        =   1
      Top             =   735
      Width           =   3465
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      ItemData        =   "Light.frx":0004
      Left            =   495
      List            =   "Light.frx":001D
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   345
      Width           =   3465
   End
   Begin VB.Label BLUZ 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   7
      Left            =   3525
      TabIndex        =   15
      Top             =   2430
      Width           =   270
   End
   Begin VB.Label BLUZ 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   6
      Left            =   3120
      TabIndex        =   14
      Top             =   2430
      Width           =   270
   End
   Begin VB.Label BLUZ 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   5
      Left            =   2700
      TabIndex        =   13
      Top             =   2430
      Width           =   270
   End
   Begin VB.Label BLUZ 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   4
      Left            =   2310
      TabIndex        =   12
      Top             =   2430
      Width           =   270
   End
   Begin VB.Label BLUZ 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   3
      Left            =   1920
      TabIndex        =   11
      Top             =   2430
      Width           =   270
   End
   Begin VB.Label BLUZ 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   2
      Left            =   1515
      TabIndex        =   10
      Top             =   2430
      Width           =   270
   End
   Begin VB.Label BLUZ 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   1
      Left            =   1110
      TabIndex        =   9
      Top             =   2430
      Width           =   270
   End
   Begin VB.Label BLUZ 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   0
      Left            =   705
      TabIndex        =   8
      Top             =   2430
      Width           =   270
   End
   Begin VB.Shape LUZ 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   225
      Index           =   7
      Left            =   3690
      Shape           =   3  'Circle
      Top             =   3075
      Width           =   255
   End
   Begin VB.Shape LUZ 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   225
      Index           =   6
      Left            =   3255
      Shape           =   3  'Circle
      Top             =   3075
      Width           =   255
   End
   Begin VB.Shape LUZ 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   225
      Index           =   5
      Left            =   2790
      Shape           =   3  'Circle
      Top             =   3075
      Width           =   255
   End
   Begin VB.Shape LUZ 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   225
      Index           =   4
      Left            =   2340
      Shape           =   3  'Circle
      Top             =   3075
      Width           =   255
   End
   Begin VB.Shape LUZ 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   225
      Index           =   3
      Left            =   1890
      Shape           =   3  'Circle
      Top             =   3075
      Width           =   255
   End
   Begin VB.Shape LUZ 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   225
      Index           =   2
      Left            =   1455
      Shape           =   3  'Circle
      Top             =   3075
      Width           =   255
   End
   Begin VB.Shape LUZ 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   225
      Index           =   1
      Left            =   990
      Shape           =   3  'Circle
      Top             =   3075
      Width           =   255
   End
   Begin VB.Label boton 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Random"
      Height          =   255
      Index           =   3
      Left            =   3090
      TabIndex        =   6
      ToolTipText     =   "Gererar Secuencia Aleatoria."
      Top             =   2685
      Width           =   735
   End
   Begin VB.Label boton 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Vaciar"
      Height          =   255
      Index           =   2
      Left            =   2295
      TabIndex        =   5
      ToolTipText     =   "Eliminar Toda La Secuencia."
      Top             =   2685
      Width           =   720
   End
   Begin VB.Label boton 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Eliminar"
      Height          =   255
      Index           =   1
      Left            =   1485
      TabIndex        =   4
      ToolTipText     =   "Suprimir Una Línea De La Secuencia."
      Top             =   2685
      Width           =   720
   End
   Begin VB.Label boton 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Agregar"
      Height          =   255
      Index           =   0
      Left            =   660
      TabIndex        =   3
      ToolTipText     =   "Agregar Una Línea A La Secuencia."
      Top             =   2670
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
      Left            =   4170
      TabIndex        =   2
      ToolTipText     =   "Minimizar"
      Top             =   45
      Width           =   225
   End
   Begin VB.Shape LUZ 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   225
      Index           =   0
      Left            =   555
      Shape           =   3  'Circle
      Top             =   3075
      Width           =   255
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   3495
      Left            =   0
      Picture         =   "Light.frx":005D
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CONTER As Integer
Dim N As Integer

Private Sub BLUZ_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 If InStr(List1.List(List1.ListIndex), Trim(str(Index + 1))) > 0 Then
  If List1.List(List1.ListIndex) <> Trim(str(Index + 1)) Then
   List1.List(List1.ListIndex) = Trim(Replace(List1.List(List1.ListIndex), Trim(str(Index + 1)), ""))
  Else
   List1.List(List1.ListIndex) = "0"
  End If
  BLUZ(Index).ForeColor = SELCOLOR
 Else
  If List1.List(List1.ListIndex) <> "0" Then
   List1.List(List1.ListIndex) = List1.List(List1.ListIndex) + Trim(str(Index + 1))
  Else
   List1.List(List1.ListIndex) = Trim(str(Index + 1))
  End If
  BLUZ(Index).ForeColor = ONCOLOR
 End If
End Sub

Private Sub BLUZ_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 N = 0
 Do While N <= 7
  If Index = BLUZ(N).Index Then
   If BLUZ(N).ForeColor <> SELCOLOR And BLUZ(N).ForeColor <> ONCOLOR Then BLUZ(N).ForeColor = SELCOLOR
  Else
   If N <= 3 Then
    If boton(N).ForeColor <> UNSELCOLOR Then boton(N).ForeColor = UNSELCOLOR
   End If
   If BLUZ(N).ForeColor <> UNSELCOLOR And BLUZ(N).ForeColor <> ONCOLOR Then BLUZ(N).ForeColor = UNSELCOLOR
  End If
  N = N + 1
 Loop
End Sub

Private Sub boton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim PAR As Boolean
 Dim pal As String
 Select Case Index
 Case Is = 0
  N = List1.ListIndex + 1
  List1.AddItem 0, List1.ListIndex + 1
  List1.ListIndex = N
 Case Is = 1
  If List1.ListIndex > 0 Then
   N = List1.ListIndex - 1
   List1.RemoveItem List1.ListIndex
  Else
   N = List1.ListIndex
   List1.RemoveItem List1.ListIndex
  End If
  If List1.ListCount > 0 Then List1.ListIndex = N
 Case Is = 2
  List1.Clear
 Case Is = 3
  List1.Clear
  N = Int(Rnd(14) * 15)
  If N = 0 Then N = 4
  PAR = False
  Do While List1.ListCount - 1 <= N
   If PAR = False Then
    List1.AddItem "0"
    PAR = True
   Else
    Select Case Int(Rnd(4) * 5)
    Case Is = 0
     pal = Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8)))
    Case Is = 1
     pal = Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8)))
    Case Is = 2
     pal = Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8)))
    Case Is = 3
     pal = Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8)))
    Case Is = 4
     pal = Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8)))
    Case Is = 5
     pal = Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8)))
    Case Is = 6
     pal = Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8)))
    Case Is = 7
     pal = Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8)))
    Case Is = 8
     pal = Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8)))
    Case Is = 9
     pal = Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8)))
    Case Is = 10
     pal = Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8)))
    Case Is = 11
     pal = Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8)))
    Case Is = 12
     pal = Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8)))
    Case Is = 13
     pal = Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8))) + Trim(str(Int(Rnd(7) * 8)))
    End Select
    List1.AddItem Replace(pal, "0", "4")
    PAR = False
   End If
  Loop
  If List1.List(List1.ListCount - 1) = "0" Then List1.RemoveItem (List1.ListCount - 1)
 End Select
End Sub

Private Sub boton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 N = 0
 Do While N <= 7
  If N <= 3 Then
   If Index = boton(N).Index Then
    If boton(N).ForeColor <> SELCOLOR Then boton(N).ForeColor = SELCOLOR
   Else
    If boton(N).ForeColor <> UNSELCOLOR Then boton(N).ForeColor = UNSELCOLOR
   End If
  End If
  If BLUZ(N).ForeColor <> UNSELCOLOR And BLUZ(N).ForeColor <> ONCOLOR Then BLUZ(N).ForeColor = UNSELCOLOR
  N = N + 1
 Loop
End Sub

Private Sub Combo1_Click()
 Select Case Combo1.ListIndex
 Case Is = 0
  List1.Clear
  List1.AddItem 0
  List1.AddItem 14
  List1.AddItem 0
  List1.AddItem 35
 End Select
End Sub

Private Sub ENC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If ENC.BackColor = &HFF& Then
  If MsgBox("ATENCIÓN: Encender el sistema puede causar daño en su hardware, sólo debe encenderlo si tiene conectado el dispositivo electrónico controlador señalado en la sección hardware de la ayuda adjunta. No nos responsabilizamos por daños causados a su hardware en caso de un mal uso de esta función. ¿Desea continuar?", vbYesNo) = vbYes Then
   ENC.BackColor = &HFFFF&
   PortOut 888, 0
   LUCES = "0"
   CONTER = 0
  End If
 Else
  ENC.BackColor = &HFF&
  PortOut 888, 0
 End If
End Sub

Private Sub Form_Load()
 Combo1.ForeColor = DATAFORE
 Combo1.BackColor = DATABACK
 List1.ForeColor = DATAFORE
 List1.BackColor = DATABACK
 N = 0
 Do While N <= 7
  BLUZ(N).ForeColor = UNSELCOLOR
  If N <= 3 Then boton(N).ForeColor = UNSELCOLOR
  N = N + 1
 Loop
 Combo1.ListIndex = 0
 List1.ListIndex = 0
 Randomize
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 CURX = X
 CURY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 0 Then
  If Label10.ForeColor = &HFFFF& Then Label10.ForeColor = &HFF&
  N = 0
  Do While N <= 7
   If BLUZ(N).ForeColor <> UNSELCOLOR And BLUZ(N).ForeColor <> ONCOLOR Then BLUZ(N).ForeColor = UNSELCOLOR
   If N <= 3 Then
    If boton(N).ForeColor <> UNSELCOLOR Then boton(N).ForeColor = UNSELCOLOR
   End If
   N = N + 1
  Loop
 Else
  Me.Move Me.left + (X - CURX), Me.top + (Y - CURY)
 End If
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Me.WindowState = 1
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Label10.ForeColor = &HFF& Then Label10.ForeColor = &HFFFF&
End Sub

Private Sub List1_Click()
 N = 0
 Do While N <= 7
  If InStr(List1.List(List1.ListIndex), Trim(str(N + 1))) > 0 Then
   If BLUZ(N).ForeColor <> ONCOLOR Then BLUZ(N).ForeColor = ONCOLOR
  Else
   If BLUZ(N).ForeColor <> UNSELCOLOR Then BLUZ(N).ForeColor = UNSELCOLOR
  End If
  N = N + 1
 Loop
End Sub

Private Sub Timer1_Timer()
  If ENC.BackColor = &HFFFF& Then
   If CONTER < List1.ListCount - 1 Then
    CONTER = CONTER + 1
   Else
    CONTER = 0
   End If
   If InStr(List1.List(CONTER), "0") > 0 Then
    If LUZ(0).BackColor <> &H0& Then LUZ(0).BackColor = &H0&
    If LUZ(1).BackColor <> &H0& Then LUZ(1).BackColor = &H0&
    If LUZ(2).BackColor <> &H0& Then LUZ(2).BackColor = &H0&
    If LUZ(3).BackColor <> &H0& Then LUZ(3).BackColor = &H0&
    If LUZ(4).BackColor <> &H0& Then LUZ(4).BackColor = &H0&
    If LUZ(5).BackColor <> &H0& Then LUZ(5).BackColor = &H0&
    If LUZ(6).BackColor <> &H0& Then LUZ(6).BackColor = &H0&
    If LUZ(7).BackColor <> &H0& Then LUZ(7).BackColor = &H0&
   Else
    If InStr(List1.List(CONTER), "1") > 0 And LUZ(0).BackColor <> &HFFFF& Then LUZ(0).BackColor = &HFFFF&
    If InStr(List1.List(CONTER), "2") > 0 And LUZ(1).BackColor <> &HFFFF& Then LUZ(1).BackColor = &HFFFF&
    If InStr(List1.List(CONTER), "3") > 0 And LUZ(2).BackColor <> &HFFFF& Then LUZ(2).BackColor = &HFFFF&
    If InStr(List1.List(CONTER), "4") > 0 And LUZ(3).BackColor <> &HFFFF& Then LUZ(3).BackColor = &HFFFF&
    If InStr(List1.List(CONTER), "5") > 0 And LUZ(4).BackColor <> &HFFFF& Then LUZ(4).BackColor = &HFFFF&
    If InStr(List1.List(CONTER), "6") > 0 And LUZ(5).BackColor <> &HFFFF& Then LUZ(5).BackColor = &HFFFF&
    If InStr(List1.List(CONTER), "7") > 0 And LUZ(6).BackColor <> &HFFFF& Then LUZ(6).BackColor = &HFFFF&
    If InStr(List1.List(CONTER), "8") > 0 And LUZ(7).BackColor <> &HFFFF& Then LUZ(7).BackColor = &HFFFF&
    If InStr(List1.List(CONTER), "1") = 0 And LUZ(0).BackColor <> &H0& Then LUZ(0).BackColor = &H0&
    If InStr(List1.List(CONTER), "2") = 0 And LUZ(1).BackColor <> &H0& Then LUZ(1).BackColor = &H0&
    If InStr(List1.List(CONTER), "3") = 0 And LUZ(2).BackColor <> &H0& Then LUZ(2).BackColor = &H0&
    If InStr(List1.List(CONTER), "4") = 0 And LUZ(3).BackColor <> &H0& Then LUZ(3).BackColor = &H0&
    If InStr(List1.List(CONTER), "5") = 0 And LUZ(4).BackColor <> &H0& Then LUZ(4).BackColor = &H0&
    If InStr(List1.List(CONTER), "6") = 0 And LUZ(5).BackColor <> &H0& Then LUZ(5).BackColor = &H0&
    If InStr(List1.List(CONTER), "7") = 0 And LUZ(6).BackColor <> &H0& Then LUZ(6).BackColor = &H0&
    If InStr(List1.List(CONTER), "8") = 0 And LUZ(7).BackColor <> &H0& Then LUZ(7).BackColor = &H0&
   End If
  End If
End Sub
