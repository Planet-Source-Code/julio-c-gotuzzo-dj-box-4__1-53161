VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Dj Box 4 - Configuración"
   ClientHeight    =   4095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5505
   Icon            =   "Conf.frx":0000
   LinkTopic       =   "Form7"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo7 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      ItemData        =   "Conf.frx":0442
      Left            =   1365
      List            =   "Conf.frx":0444
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   3015
      Width           =   3240
   End
   Begin VB.ComboBox Combo6 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      ItemData        =   "Conf.frx":0446
      Left            =   2025
      List            =   "Conf.frx":045F
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   2550
      Width           =   2580
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   4275
      ScaleHeight     =   300
      ScaleWidth      =   330
      TabIndex        =   14
      Top             =   1995
      Width           =   330
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   4275
      ScaleHeight     =   300
      ScaleWidth      =   330
      TabIndex        =   13
      Top             =   1680
      Width           =   330
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   4275
      ScaleHeight     =   300
      ScaleWidth      =   330
      TabIndex        =   12
      Top             =   1365
      Width           =   330
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   4275
      ScaleHeight     =   300
      ScaleWidth      =   330
      TabIndex        =   11
      Top             =   1050
      Width           =   330
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   4275
      ScaleHeight     =   300
      ScaleWidth      =   330
      TabIndex        =   10
      Top             =   735
      Width           =   330
   End
   Begin VB.ComboBox Combo5 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      ItemData        =   "Conf.frx":04AD
      Left            =   2775
      List            =   "Conf.frx":04F0
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1995
      Width           =   1440
   End
   Begin VB.ComboBox Combo4 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      ItemData        =   "Conf.frx":05BB
      Left            =   2775
      List            =   "Conf.frx":05FE
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1680
      Width           =   1440
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      ItemData        =   "Conf.frx":06C9
      Left            =   2775
      List            =   "Conf.frx":070C
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1365
      Width           =   1440
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      ItemData        =   "Conf.frx":07D7
      Left            =   2775
      List            =   "Conf.frx":081A
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1050
      Width           =   1440
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      ItemData        =   "Conf.frx":08E5
      Left            =   2775
      List            =   "Conf.frx":0928
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   735
      Width           =   1440
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dj Box 4.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2160
      TabIndex        =   21
      Top             =   240
      Width           =   915
   End
   Begin VB.Label Command2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancelar"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4200
      TabIndex        =   20
      Top             =   3600
      Width           =   945
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Aceptar"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3240
      TabIndex        =   19
      Top             =   3600
      Width           =   945
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Skin:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   885
      TabIndex        =   17
      Top             =   3075
      Width           =   435
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color Themes:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   870
      TabIndex        =   15
      Top             =   2610
      Width           =   1020
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color Botón Encendido:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   855
      TabIndex        =   4
      Top             =   2040
      Width           =   1680
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color Botón Selecionado:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   855
      TabIndex        =   3
      Top             =   1725
      Width           =   1800
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color Botón General:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   855
      TabIndex        =   2
      Top             =   1410
      Width           =   1470
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color Fondo General:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   855
      TabIndex        =   1
      Top             =   1080
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color Texto General:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   855
      TabIndex        =   0
      Top             =   780
      Width           =   1455
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   4095
      Left            =   0
      Picture         =   "Conf.frx":09F3
      Top             =   0
      Width           =   5505
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim N As Integer
Dim TEMP1 As Integer
Dim TEMP2(4)

Private Sub Combo1_Click()
 Select Case Combo1.ListIndex
 Case Is = 0
  Picture1.BackColor = 16777215
 Case Is = 1
  Picture1.BackColor = 12632256
 Case Is = 2
  Picture1.BackColor = 8421504
 Case Is = 3
  Picture1.BackColor = 4210752
 Case Is = 4
  Picture1.BackColor = 0
 Case Is = 5
  Picture1.BackColor = 16711680
 Case Is = 6
  Picture1.BackColor = 8388608
 Case Is = 7
  Picture1.BackColor = 65280
 Case Is = 8
  Picture1.BackColor = 32768
 Case Is = 9
  Picture1.BackColor = 255
 Case Is = 10
  Picture1.BackColor = 192
 Case Is = 11
  Picture1.BackColor = 128
 Case Is = 12
  Picture1.BackColor = 33023
 Case Is = 13
  Picture1.BackColor = 65535
 Case Is = 14
  Picture1.BackColor = 16776960
 Case Is = 15
  Picture1.BackColor = 16711935
 Case Is = 16
  Picture1.BackColor = 12648447
 Case Is = 17
  Picture1.BackColor = 12632319
 Case Is = 18
  Picture1.BackColor = 8421376
 Case Is = 19
  Picture1.BackColor = 49344
 Case Is = 20
  Picture1.BackColor = 4194304
 End Select
 SELCOLOR = Val(Picture4.BackColor)
 UNSELCOLOR = Val(Picture3.BackColor)
 ONCOLOR = Val(Picture5.BackColor)
 DATAFORE = Val(Picture1.BackColor)
 DATABACK = Val(Picture2.BackColor)
 Call SETCOLORS
End Sub

Private Sub Combo2_Click()
 Select Case Combo2.ListIndex
 Case Is = 0
  Picture2.BackColor = 16777215
 Case Is = 1
  Picture2.BackColor = 12632256
 Case Is = 2
  Picture2.BackColor = 8421504
 Case Is = 3
  Picture2.BackColor = 4210752
 Case Is = 4
  Picture2.BackColor = 0
 Case Is = 5
  Picture2.BackColor = 16711680
 Case Is = 6
  Picture2.BackColor = 8388608
 Case Is = 7
  Picture2.BackColor = 65280
 Case Is = 8
  Picture2.BackColor = 32768
 Case Is = 9
  Picture2.BackColor = 255
 Case Is = 10
  Picture2.BackColor = 192
 Case Is = 11
  Picture2.BackColor = 128
 Case Is = 12
  Picture2.BackColor = 33023
 Case Is = 13
  Picture2.BackColor = 65535
 Case Is = 14
  Picture2.BackColor = 16776960
 Case Is = 15
  Picture2.BackColor = 16711935
 Case Is = 16
  Picture2.BackColor = 12648447
 Case Is = 17
  Picture2.BackColor = 12632319
 Case Is = 18
  Picture2.BackColor = 8421376
 Case Is = 19
  Picture2.BackColor = 49344
 Case Is = 20
  Picture2.BackColor = 4194304
 End Select
 SELCOLOR = Val(Picture4.BackColor)
 UNSELCOLOR = Val(Picture3.BackColor)
 ONCOLOR = Val(Picture5.BackColor)
 DATAFORE = Val(Picture1.BackColor)
 DATABACK = Val(Picture2.BackColor)
 Call SETCOLORS
End Sub

Private Sub Combo3_Click()
 Select Case Combo3.ListIndex
 Case Is = 0
  Picture3.BackColor = 16777215
 Case Is = 1
  Picture3.BackColor = 12632256
 Case Is = 2
  Picture3.BackColor = 8421504
 Case Is = 3
  Picture3.BackColor = 4210752
 Case Is = 4
  Picture3.BackColor = 0
 Case Is = 5
  Picture3.BackColor = 16711680
 Case Is = 6
  Picture3.BackColor = 8388608
 Case Is = 7
  Picture3.BackColor = 65280
 Case Is = 8
  Picture3.BackColor = 32768
 Case Is = 9
  Picture3.BackColor = 255
 Case Is = 10
  Picture3.BackColor = 192
 Case Is = 11
  Picture3.BackColor = 128
 Case Is = 12
  Picture3.BackColor = 33023
 Case Is = 13
  Picture3.BackColor = 65535
 Case Is = 14
  Picture3.BackColor = 16776960
 Case Is = 15
  Picture3.BackColor = 16711935
 Case Is = 16
  Picture3.BackColor = 12648447
 Case Is = 17
  Picture3.BackColor = 12632319
 Case Is = 18
  Picture3.BackColor = 8421376
 Case Is = 19
  Picture3.BackColor = 49344
 Case Is = 20
  Picture3.BackColor = 4194304
 End Select
 SELCOLOR = Val(Picture4.BackColor)
 UNSELCOLOR = Val(Picture3.BackColor)
 ONCOLOR = Val(Picture5.BackColor)
 DATAFORE = Val(Picture1.BackColor)
 DATABACK = Val(Picture2.BackColor)
 Call SETCOLORS
End Sub

Private Sub Combo4_Click()
 Select Case Combo4.ListIndex
 Case Is = 0
  Picture4.BackColor = 16777215
 Case Is = 1
  Picture4.BackColor = 12632256
 Case Is = 2
  Picture4.BackColor = 8421504
 Case Is = 3
  Picture4.BackColor = 4210752
 Case Is = 4
  Picture4.BackColor = 0
 Case Is = 5
  Picture4.BackColor = 16711680
 Case Is = 6
  Picture4.BackColor = 8388608
 Case Is = 7
  Picture4.BackColor = 65280
 Case Is = 8
  Picture4.BackColor = 32768
 Case Is = 9
  Picture4.BackColor = 255
 Case Is = 10
  Picture4.BackColor = 192
 Case Is = 11
  Picture4.BackColor = 128
 Case Is = 12
  Picture4.BackColor = 33023
 Case Is = 13
  Picture4.BackColor = 65535
 Case Is = 14
  Picture4.BackColor = 16776960
 Case Is = 15
  Picture4.BackColor = 16711935
 Case Is = 16
  Picture4.BackColor = 12648447
 Case Is = 17
  Picture4.BackColor = 12632319
 Case Is = 18
  Picture4.BackColor = 8421376
 Case Is = 19
  Picture4.BackColor = 49344
 Case Is = 20
  Picture4.BackColor = 4194304
 End Select
 SELCOLOR = Val(Picture4.BackColor)
 UNSELCOLOR = Val(Picture3.BackColor)
 ONCOLOR = Val(Picture5.BackColor)
 DATAFORE = Val(Picture1.BackColor)
 DATABACK = Val(Picture2.BackColor)
 Call SETCOLORS
End Sub

Private Sub Combo5_Click()
 Select Case Combo5.ListIndex
 Case Is = 0
  Picture5.BackColor = 16777215
 Case Is = 1
  Picture5.BackColor = 12632256
 Case Is = 2
  Picture5.BackColor = 8421504
 Case Is = 3
  Picture5.BackColor = 4210752
 Case Is = 4
  Picture5.BackColor = 0
 Case Is = 5
  Picture5.BackColor = 16711680
 Case Is = 6
  Picture5.BackColor = 8388608
 Case Is = 7
  Picture5.BackColor = 65280
 Case Is = 8
  Picture5.BackColor = 32768
 Case Is = 9
  Picture5.BackColor = 255
 Case Is = 10
  Picture5.BackColor = 192
 Case Is = 11
  Picture5.BackColor = 128
 Case Is = 12
  Picture5.BackColor = 33023
 Case Is = 13
  Picture5.BackColor = 65535
 Case Is = 14
  Picture5.BackColor = 16776960
 Case Is = 15
  Picture5.BackColor = 16711935
 Case Is = 16
  Picture5.BackColor = 12648447
 Case Is = 17
  Picture5.BackColor = 12632319
 Case Is = 18
  Picture5.BackColor = 8421376
 Case Is = 19
  Picture5.BackColor = 49344
 Case Is = 20
  Picture5.BackColor = 4194304
 End Select
 SELCOLOR = Val(Picture4.BackColor)
 UNSELCOLOR = Val(Picture3.BackColor)
 ONCOLOR = Val(Picture5.BackColor)
 DATAFORE = Val(Picture1.BackColor)
 DATABACK = Val(Picture2.BackColor)
 Call SETCOLORS
End Sub

Private Sub Combo6_Click()
 Select Case Combo6.ListIndex
 Case Is = 0
  Combo1.ListIndex = 7
  Combo2.ListIndex = 4
  Combo3.ListIndex = 4
  Combo4.ListIndex = 13
  Combo5.ListIndex = 5
 Case Is = 1
  Combo1.ListIndex = 4
  Combo2.ListIndex = 7
  Combo3.ListIndex = 4
  Combo4.ListIndex = 14
  Combo5.ListIndex = 9
 Case Is = 2
  Combo1.ListIndex = 0
  Combo2.ListIndex = 15
  Combo3.ListIndex = 11
  Combo4.ListIndex = 17
  Combo5.ListIndex = 12
 Case Is = 3
  Combo1.ListIndex = 9
  Combo2.ListIndex = 4
  Combo3.ListIndex = 11
  Combo4.ListIndex = 12
  Combo5.ListIndex = 13
 Case Is = 4
  Combo1.ListIndex = 0
  Combo2.ListIndex = 4
  Combo3.ListIndex = 4
  Combo4.ListIndex = 3
  Combo5.ListIndex = 1
 Case Is = 5
  Combo1.ListIndex = 4
  Combo2.ListIndex = 0
  Combo3.ListIndex = 4
  Combo4.ListIndex = 1
  Combo5.ListIndex = 0
 Case Is = 6
  Combo1.ListIndex = 12
  Combo2.ListIndex = 20
  Combo3.ListIndex = 20
  Combo4.ListIndex = 5
  Combo5.ListIndex = 19
 End Select
 SELCOLOR = Val(Picture4.BackColor)
 UNSELCOLOR = Val(Picture3.BackColor)
 ONCOLOR = Val(Picture5.BackColor)
 DATAFORE = Val(Picture1.BackColor)
 DATABACK = Val(Picture2.BackColor)
 Call SETCOLORS
End Sub

Private Sub Combo7_Click()
 If Combo7.ListCount > 0 Then
  If Combo7.ListIndex = 0 Then
   SKINFILE = 0
  Else
   If SKINFILE <> Combo7.ListIndex Then
    SKINFILE = Combo7.ListIndex
    Call SET_SKIN(0, SKINFILE)
    Call SET_SKIN(1, SKINFILE)
    Call SET_SKIN(2, SKINFILE)
    Call SET_SKIN(3, SKINFILE)
   End If
  End If
 End If
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Form1.Enabled = True
 Unload Me
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Command2.ForeColor <> UNSELCOLOR Then Command2.ForeColor = UNSELCOLOR
 If Command1.ForeColor = UNSELCOLOR Then Command1.ForeColor = SELCOLOR
End Sub

Private Sub Command2_Click()
 DATAFORE = TEMP2(0)
 DATABACK = TEMP2(1)
 UNSELCOLOR = TEMP2(2)
 SELCOLOR = TEMP2(3)
 ONCOLOR = TEMP2(4)
 Call SETCOLORS
 SKINFILE = TEMP1
 If SKINFILE > 0 Then
  Call SET_SKIN(0, SKINFILE)
  Call SET_SKIN(1, SKINFILE)
  Call SET_SKIN(2, SKINFILE)
  Call SET_SKIN(3, SKINFILE)
 End If
 Form1.Enabled = True
 Unload Me
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Command1.ForeColor <> UNSELCOLOR Then Command1.ForeColor = UNSELCOLOR
 If Command2.ForeColor = UNSELCOLOR Then Command2.ForeColor = SELCOLOR
End Sub

Private Sub Form_Load()
 TEMP1 = SKINFILE
 TEMP2(0) = DATAFORE
 TEMP2(1) = DATABACK
 TEMP2(2) = UNSELCOLOR
 TEMP2(3) = SELCOLOR
 TEMP2(4) = ONCOLOR
 Select Case DATAFORE
 Case Is = 16777215
  Combo1.ListIndex = 0
 Case Is = 12632256
  Combo1.ListIndex = 1
 Case Is = 8421504
  Combo1.ListIndex = 2
 Case Is = 4210752
  Combo1.ListIndex = 3
 Case Is = 0
  Combo1.ListIndex = 4
 Case Is = 16711680
  Combo1.ListIndex = 5
 Case Is = 8388608
  Combo1.ListIndex = 6
 Case Is = 65280
  Combo1.ListIndex = 7
 Case Is = 32768
  Combo1.ListIndex = 8
 Case Is = 255
  Combo1.ListIndex = 9
 Case Is = 192
  Combo1.ListIndex = 10
 Case Is = 128
  Combo1.ListIndex = 11
 Case Is = 33023
  Combo1.ListIndex = 12
 Case Is = 65535
  Combo1.ListIndex = 13
 Case Is = 16776960
  Combo1.ListIndex = 14
 Case Is = 16711935
  Combo1.ListIndex = 15
 Case Is = 12648447
  Combo1.ListIndex = 16
 Case Is = 12632319
  Combo1.ListIndex = 17
 Case Is = 8421376
  Combo1.ListIndex = 18
 Case Is = 49344
  Combo1.ListIndex = 19
 Case Is = 4194304
  Combo1.ListIndex = 20
 End Select
 Select Case DATABACK
 Case Is = 16777215
  Combo2.ListIndex = 0
 Case Is = 12632256
  Combo2.ListIndex = 1
 Case Is = 8421504
  Combo2.ListIndex = 2
 Case Is = 4210752
  Combo2.ListIndex = 3
 Case Is = 0
  Combo2.ListIndex = 4
 Case Is = 16711680
  Combo2.ListIndex = 5
 Case Is = 8388608
  Combo2.ListIndex = 6
 Case Is = 65280
  Combo2.ListIndex = 7
 Case Is = 32768
  Combo2.ListIndex = 8
 Case Is = 255
  Combo2.ListIndex = 9
 Case Is = 192
  Combo2.ListIndex = 10
 Case Is = 128
  Combo2.ListIndex = 11
 Case Is = 33023
  Combo2.ListIndex = 12
 Case Is = 65535
  Combo2.ListIndex = 13
 Case Is = 16776960
  Combo2.ListIndex = 14
 Case Is = 16711935
  Combo2.ListIndex = 15
 Case Is = 12648447
  Combo2.ListIndex = 16
 Case Is = 12632319
  Combo2.ListIndex = 17
 Case Is = 8421376
  Combo2.ListIndex = 18
 Case Is = 49344
  Combo2.ListIndex = 19
 Case Is = 4194304
  Combo2.ListIndex = 20
 End Select
 Select Case UNSELCOLOR
 Case Is = 16777215
  Combo3.ListIndex = 0
 Case Is = 12632256
  Combo3.ListIndex = 1
 Case Is = 8421504
  Combo3.ListIndex = 2
 Case Is = 4210752
  Combo3.ListIndex = 3
 Case Is = 0
  Combo3.ListIndex = 4
 Case Is = 16711680
  Combo3.ListIndex = 5
 Case Is = 8388608
  Combo3.ListIndex = 6
 Case Is = 65280
  Combo3.ListIndex = 7
 Case Is = 32768
  Combo3.ListIndex = 8
 Case Is = 255
  Combo3.ListIndex = 9
 Case Is = 192
  Combo3.ListIndex = 10
 Case Is = 128
  Combo3.ListIndex = 11
 Case Is = 33023
  Combo3.ListIndex = 12
 Case Is = 65535
  Combo3.ListIndex = 13
 Case Is = 16776960
  Combo3.ListIndex = 14
 Case Is = 16711935
  Combo3.ListIndex = 15
 Case Is = 12648447
  Combo3.ListIndex = 16
 Case Is = 12632319
  Combo3.ListIndex = 17
 Case Is = 8421376
  Combo3.ListIndex = 18
 Case Is = 49344
  Combo3.ListIndex = 19
 Case Is = 4194304
  Combo3.ListIndex = 20
 End Select
 Select Case SELCOLOR
 Case Is = 16777215
  Combo4.ListIndex = 0
 Case Is = 12632256
  Combo4.ListIndex = 1
 Case Is = 8421504
  Combo4.ListIndex = 2
 Case Is = 4210752
  Combo4.ListIndex = 3
 Case Is = 0
  Combo4.ListIndex = 4
 Case Is = 16711680
  Combo4.ListIndex = 5
 Case Is = 8388608
  Combo4.ListIndex = 6
 Case Is = 65280
  Combo4.ListIndex = 7
 Case Is = 32768
  Combo4.ListIndex = 8
 Case Is = 255
  Combo4.ListIndex = 9
 Case Is = 192
  Combo4.ListIndex = 10
 Case Is = 128
  Combo4.ListIndex = 11
 Case Is = 33023
  Combo4.ListIndex = 12
 Case Is = 65535
  Combo4.ListIndex = 13
 Case Is = 16776960
  Combo4.ListIndex = 14
 Case Is = 16711935
  Combo4.ListIndex = 15
 Case Is = 12648447
  Combo4.ListIndex = 16
 Case Is = 12632319
  Combo4.ListIndex = 17
 Case Is = 8421376
  Combo4.ListIndex = 18
 Case Is = 49344
  Combo4.ListIndex = 19
 Case Is = 4194304
  Combo4.ListIndex = 20
 End Select
 Select Case ONCOLOR
 Case Is = 16777215
  Combo5.ListIndex = 0
 Case Is = 12632256
  Combo5.ListIndex = 1
 Case Is = 8421504
  Combo5.ListIndex = 2
 Case Is = 4210752
  Combo5.ListIndex = 3
 Case Is = 0
  Combo5.ListIndex = 4
 Case Is = 16711680
  Combo5.ListIndex = 5
 Case Is = 8388608
  Combo5.ListIndex = 6
 Case Is = 65280
  Combo5.ListIndex = 7
 Case Is = 32768
  Combo5.ListIndex = 8
 Case Is = 255
  Combo5.ListIndex = 9
 Case Is = 192
  Combo5.ListIndex = 10
 Case Is = 128
  Combo5.ListIndex = 11
 Case Is = 33023
  Combo5.ListIndex = 12
 Case Is = 65535
  Combo5.ListIndex = 13
 Case Is = 16776960
  Combo5.ListIndex = 14
 Case Is = 16711935
  Combo5.ListIndex = 15
 Case Is = 12648447
  Combo5.ListIndex = 16
 Case Is = 12632319
  Combo5.ListIndex = 17
 Case Is = 8421376
  Combo5.ListIndex = 18
 Case Is = 49344
  Combo5.ListIndex = 19
 Case Is = 4194304
  Combo5.ListIndex = 20
 End Select
 If Dir(App.path & "\Skins\skins.dat", vbArchive) <> "" Then
  If SKINS_LIST(Combo7) = False Then
   MsgBox ("ERROR: Archivo Skins.dat dañado, imposible continuar con la carga.")
  Else
   Combo7.ListIndex = SKINFILE
  End If
 End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Command1.ForeColor <> UNSELCOLOR Then Command1.ForeColor = UNSELCOLOR
 If Command2.ForeColor <> UNSELCOLOR Then Command2.ForeColor = UNSELCOLOR
End Sub

Public Sub SETCOLORS()
  Dim N As Integer
  If Form1.LISTA(0).ForeColor <> DATAFORE Then Form1.LISTA(0).ForeColor = DATAFORE
  If Form1.LISTA(0).BackColor <> DATABACK Then Form1.LISTA(0).BackColor = DATABACK
  If Form1.LISTB(0).ForeColor <> DATAFORE Then Form1.LISTB(0).ForeColor = DATAFORE
  If Form1.LISTB(0).BackColor <> DATABACK Then Form1.LISTB(0).BackColor = DATABACK
  If Form1.INICIOA.ForeColor <> DATAFORE Then Form1.INICIOA.ForeColor = DATAFORE
  If Form1.INICIOA.BackColor <> DATABACK Then Form1.INICIOA.BackColor = DATABACK
  If Form1.INICIOB.ForeColor <> DATAFORE Then Form1.INICIOB.ForeColor = DATAFORE
  If Form1.INICIOB.BackColor <> DATABACK Then Form1.INICIOB.BackColor = DATABACK
  If Form1.FINA.ForeColor <> DATAFORE Then Form1.FINA.ForeColor = DATAFORE
  If Form1.FINA.BackColor <> DATABACK Then Form1.FINA.BackColor = DATABACK
  If Form1.FINB.ForeColor <> DATAFORE Then Form1.FINB.ForeColor = DATAFORE
  If Form1.FINB.BackColor <> DATABACK Then Form1.FINB.BackColor = DATABACK
  If Form1.NUMA.ForeColor <> DATAFORE Then Form1.NUMA.ForeColor = DATAFORE
  If Form1.NUMB.ForeColor <> DATAFORE Then Form1.NUMB.ForeColor = DATAFORE
  If Form1.TITLEA.ForeColor <> DATAFORE Then Form1.TITLEA.ForeColor = DATAFORE
  If Form1.TITLEB.ForeColor <> DATAFORE Then Form1.TITLEB.ForeColor = DATAFORE
  If Form1.CURPOSA.ForeColor <> DATAFORE Then Form1.CURPOSA.ForeColor = DATAFORE
  If Form1.CURPOSB.ForeColor <> DATAFORE Then Form1.CURPOSB.ForeColor = DATAFORE
  If Form1.DURA.ForeColor <> DATAFORE Then Form1.DURA.ForeColor = DATAFORE
  If Form1.DURB.ForeColor <> DATAFORE Then Form1.DURB.ForeColor = DATAFORE
  If Form1.Picture1.BackColor <> DATABACK Then Form1.Picture1.BackColor = DATABACK
  If Form1.Picture2.BackColor <> DATABACK Then Form1.Picture2.BackColor = DATABACK
  If Form8.Picture1.BackColor <> DATABACK Then Form8.Picture1.BackColor = DATABACK
  If Form8.Combo1.BackColor <> DATABACK Then Form8.Combo1.BackColor = DATABACK
  If Form8.Combo1.ForeColor <> DATAFORE Then Form8.Combo1.ForeColor = DATAFORE
  If Form8.Label2.ForeColor <> DATAFORE Then Form8.Label2.ForeColor = DATAFORE
  If Form8.Label3.ForeColor <> DATAFORE Then Form8.Label3.ForeColor = DATAFORE
  If Form8.Label4.ForeColor <> DATAFORE Then Form8.Label4.ForeColor = DATAFORE
  If Form8.Label5.ForeColor <> DATAFORE Then Form8.Label5.ForeColor = DATAFORE
  If Form10.RANBUT.ForeColor <> UNSELCOLOR Then Form10.RANBUT.ForeColor = UNSELCOLOR
  If Form10.LOOPBUT.ForeColor <> UNSELCOLOR Then Form10.LOOPBUT.ForeColor = UNSELCOLOR
  If Form10.PANBUT.ForeColor <> UNSELCOLOR Then Form10.PANBUT.ForeColor = UNSELCOLOR
  N = 0
  Do While N <= 15
   If Form1.COMANDO(N).ForeColor <> UNSELCOLOR Then Form1.COMANDO(N).ForeColor = UNSELCOLOR
   If Form4.FX(N).ForeColor <> UNSELCOLOR Then Form4.FX(N).ForeColor = UNSELCOLOR
   If Form4.RTM(N).ForeColor <> UNSELCOLOR Then Form4.RTM(N).ForeColor = UNSELCOLOR
   If N <= 3 Then
    If Form1.MARCA(N).ForeColor <> UNSELCOLOR Then Form1.MARCA(N).ForeColor = UNSELCOLOR
    If Form10.Command1(N).ForeColor <> UNSELCOLOR Then Form10.Command1(N).ForeColor = UNSELCOLOR
   End If
   If N <= 4 Then
    If Form8.Label1(N).ForeColor <> UNSELCOLOR Then Form8.Label1(N).ForeColor = UNSELCOLOR
   End If
   N = N + 1
  Loop
  N = 0
  If Form4.RTM(0).ForeColor <> ONCOLOR Then Form4.RTM(0).ForeColor = ONCOLOR
End Sub
