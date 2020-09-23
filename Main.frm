VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Dj Box 4 - Player"
   ClientHeight    =   4665
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   7455
   ForeColor       =   &H00000000&
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3270
      TabIndex        =   58
      Text            =   "Text2"
      Top             =   15
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   855
      TabIndex        =   57
      Text            =   "Text1"
      Top             =   45
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.Timer TIMERFX 
      Interval        =   200
      Left            =   240
      Top             =   3360
   End
   Begin PicClip.PictureClip STATECLIP 
      Left            =   -630
      Top             =   3495
      _ExtentX        =   1984
      _ExtentY        =   609
      _Version        =   393216
      Cols            =   3
      Picture         =   "Main.frx":0442
   End
   Begin MSComDlg.CommonDialog CM 
      Left            =   0
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox LISTA 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H0000FF00&
      Height          =   1230
      Index           =   1
      ItemData        =   "Main.frx":086C
      Left            =   225
      List            =   "Main.frx":086E
      TabIndex        =   36
      Top             =   1275
      Visible         =   0   'False
      Width           =   315
   End
   Begin ComctlLib.Slider VEL 
      Height          =   645
      Index           =   0
      Left            =   540
      TabIndex        =   45
      Top             =   240
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   1138
      _Version        =   327682
      Orientation     =   1
      Max             =   40
      SelStart        =   20
      TickStyle       =   3
      Value           =   20
   End
   Begin ComctlLib.Slider TIMEA 
      Height          =   255
      Left            =   870
      TabIndex        =   43
      Top             =   945
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   450
      _Version        =   327682
      TickStyle       =   3
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   150
      Top             =   4140
   End
   Begin VB.ListBox FLISTB 
      Height          =   1035
      ItemData        =   "Main.frx":0870
      Left            =   6720
      List            =   "Main.frx":0872
      TabIndex        =   41
      Top             =   4440
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.ListBox ILISTB 
      Height          =   1035
      ItemData        =   "Main.frx":0874
      Left            =   6960
      List            =   "Main.frx":0876
      TabIndex        =   40
      Top             =   4440
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.ListBox FLISTA 
      Height          =   1035
      ItemData        =   "Main.frx":0878
      Left            =   -1080
      List            =   "Main.frx":087A
      TabIndex        =   39
      Top             =   4440
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.ListBox ILISTA 
      ForeColor       =   &H00FF0000&
      Height          =   1425
      ItemData        =   "Main.frx":087C
      Left            =   -600
      List            =   "Main.frx":087E
      TabIndex        =   38
      Top             =   4440
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.ListBox LISTB 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H0000FF00&
      Height          =   840
      Index           =   1
      ItemData        =   "Main.frx":0880
      Left            =   6960
      List            =   "Main.frx":0882
      TabIndex        =   37
      Top             =   2760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox FINB 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   5625
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "00:00"
      Top             =   2010
      Width           =   645
   End
   Begin VB.TextBox INICIOB 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   4275
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "00:00"
      Top             =   2025
      Width           =   645
   End
   Begin VB.TextBox FINA 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   2535
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "00:00"
      Top             =   2010
      Width           =   645
   End
   Begin VB.TextBox INICIOA 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   1185
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "00:00"
      Top             =   2010
      Width           =   630
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      ForeColor       =   &H00000000&
      Height          =   660
      Left            =   3915
      ScaleHeight     =   600
      ScaleWidth      =   2715
      TabIndex        =   3
      Top             =   225
      Width           =   2775
      Begin VB.Label ACTION 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   180
         Index           =   1
         Left            =   2355
         TabIndex        =   56
         Top             =   390
         Width           =   330
      End
      Begin VB.Image STATEBMP 
         Height          =   150
         Index           =   1
         Left            =   30
         Stretch         =   -1  'True
         Top             =   15
         Width           =   165
      End
      Begin VB.Label DURB 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "- | 00:00 |"
         ForeColor       =   &H0000FF00&
         Height          =   225
         Left            =   1530
         TabIndex        =   15
         Top             =   285
         Width           =   900
      End
      Begin VB.Label CURPOSB 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "< 00:00 >"
         ForeColor       =   &H0000FF00&
         Height          =   225
         Left            =   855
         TabIndex        =   14
         Top             =   285
         Width           =   765
      End
      Begin VB.Label TITLEB 
         BackStyle       =   0  'Transparent
         Caption         =   """ Vacío """
         ForeColor       =   &H0000FF00&
         Height          =   225
         Left            =   795
         TabIndex        =   13
         Top             =   30
         Width           =   1755
      End
      Begin VB.Label NUMB 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   450
         Left            =   60
         TabIndex        =   9
         Top             =   105
         Width           =   675
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00000000&
      Height          =   660
      Left            =   795
      ScaleHeight     =   600
      ScaleWidth      =   2715
      TabIndex        =   2
      Top             =   225
      Width           =   2775
      Begin VB.Label ACTION 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   180
         Index           =   0
         Left            =   2355
         TabIndex        =   55
         Top             =   390
         Width           =   330
      End
      Begin VB.Image STATEBMP 
         Height          =   150
         Index           =   0
         Left            =   30
         Stretch         =   -1  'True
         Top             =   15
         Width           =   165
      End
      Begin VB.Label DURA 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "- | 00:00 |"
         ForeColor       =   &H0000FF00&
         Height          =   225
         Left            =   1515
         TabIndex        =   12
         Top             =   300
         Width           =   900
      End
      Begin VB.Label CURPOSA 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "< 00:00 >"
         ForeColor       =   &H0000FF00&
         Height          =   225
         Left            =   855
         TabIndex        =   11
         Top             =   300
         Width           =   765
      End
      Begin VB.Label TITLEA 
         BackStyle       =   0  'Transparent
         Caption         =   """ Vacío """
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   780
         TabIndex        =   10
         ToolTipText     =   "1"
         Top             =   60
         Width           =   1755
      End
      Begin VB.Label NUMA 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   435
         Left            =   60
         TabIndex        =   8
         Top             =   105
         Width           =   675
      End
   End
   Begin VB.ListBox LISTB 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   2010
      Index           =   0
      ItemData        =   "Main.frx":0884
      Left            =   3900
      List            =   "Main.frx":0886
      TabIndex        =   1
      Top             =   2295
      Width           =   2775
   End
   Begin VB.ListBox LISTA 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   2010
      Index           =   0
      ItemData        =   "Main.frx":0888
      Left            =   795
      List            =   "Main.frx":088A
      TabIndex        =   0
      Top             =   2295
      Width           =   2775
   End
   Begin ComctlLib.Slider TIMEB 
      Height          =   255
      Left            =   3990
      TabIndex        =   44
      Top             =   945
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   450
      _Version        =   327682
      TickStyle       =   3
   End
   Begin ComctlLib.Slider VEL 
      Height          =   645
      Index           =   1
      Left            =   6735
      TabIndex        =   46
      Top             =   225
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   1138
      _Version        =   327682
      Orientation     =   1
      Max             =   40
      SelStart        =   20
      TickStyle       =   3
      Value           =   20
   End
   Begin VB.Label Command6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "m3u List"
      Height          =   225
      Left            =   5835
      TabIndex        =   54
      Top             =   4335
      Width           =   840
   End
   Begin VB.Label Command5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Vaciar"
      Height          =   225
      Left            =   4875
      TabIndex        =   53
      Top             =   4335
      Width           =   840
   End
   Begin VB.Label Command2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar"
      Height          =   225
      Left            =   3915
      TabIndex        =   52
      Top             =   4335
      Width           =   840
   End
   Begin VB.Label Command4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "m3u List"
      Height          =   225
      Left            =   2730
      TabIndex        =   51
      Top             =   4335
      Width           =   840
   End
   Begin VB.Label Command3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Vaciar"
      Height          =   225
      Left            =   1785
      TabIndex        =   50
      Top             =   4335
      Width           =   840
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar"
      Height          =   225
      Left            =   810
      TabIndex        =   49
      Top             =   4335
      Width           =   840
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   60
      TabIndex        =   48
      ToolTipText     =   "Configuración"
      Top             =   105
      Width           =   210
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6945
      TabIndex        =   47
      ToolTipText     =   "Minimizar Todo"
      Top             =   75
      Width           =   255
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   165
      Left            =   7215
      TabIndex        =   42
      ToolTipText     =   "Cerrar"
      Top             =   120
      Width           =   150
   End
   Begin VB.Label MARCA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      Height          =   240
      Index           =   3
      Left            =   6345
      TabIndex        =   35
      ToolTipText     =   "Marcar El Final Del Tema."
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label MARCA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      Height          =   240
      Index           =   2
      Left            =   3945
      TabIndex        =   34
      ToolTipText     =   "Marcar El Inicio Del Tema."
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label MARCA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      Height          =   240
      Index           =   1
      Left            =   3255
      TabIndex        =   33
      ToolTipText     =   "Marcar El Final Del Tema."
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label MARCA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   870
      TabIndex        =   32
      ToolTipText     =   "Marcar El Inicio Del Tema."
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label COMANDO 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Play"
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   15
      Left            =   3930
      TabIndex        =   31
      ToolTipText     =   "Reproducir."
      Top             =   1305
      Width           =   600
   End
   Begin VB.Label COMANDO 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pause"
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   14
      Left            =   4620
      TabIndex        =   30
      ToolTipText     =   "Pausar."
      Top             =   1305
      Width           =   600
   End
   Begin VB.Label COMANDO 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Prev."
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   13
      Left            =   5370
      TabIndex        =   29
      ToolTipText     =   "Tema Anterior."
      Top             =   1305
      Width           =   600
   End
   Begin VB.Label COMANDO 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Next"
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   12
      Left            =   6075
      TabIndex        =   28
      ToolTipText     =   "Tema Siguiente."
      Top             =   1305
      Width           =   600
   End
   Begin VB.Label COMANDO 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Rew."
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   11
      Left            =   3945
      TabIndex        =   27
      ToolTipText     =   "Rebobinar"
      Top             =   1680
      Width           =   600
   End
   Begin VB.Label COMANDO 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FF."
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   10
      Left            =   4665
      TabIndex        =   26
      ToolTipText     =   "Adelantar."
      Top             =   1680
      Width           =   600
   End
   Begin VB.Label COMANDO 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Stop"
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   9
      Left            =   5355
      TabIndex        =   25
      ToolTipText     =   "Detener."
      Top             =   1665
      Width           =   600
   End
   Begin VB.Label COMANDO 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Open"
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   8
      Left            =   6060
      TabIndex        =   24
      ToolTipText     =   "Cargar Temas."
      Top             =   1680
      Width           =   600
   End
   Begin VB.Label COMANDO 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Open"
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   7
      Left            =   2955
      TabIndex        =   23
      ToolTipText     =   "Cargar Temas."
      Top             =   1680
      Width           =   600
   End
   Begin VB.Label COMANDO 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Stop"
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   6
      Left            =   2250
      TabIndex        =   22
      ToolTipText     =   "Detener."
      Top             =   1665
      Width           =   600
   End
   Begin VB.Label COMANDO 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FF."
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   5
      Left            =   1545
      TabIndex        =   21
      ToolTipText     =   "Adelantar."
      Top             =   1680
      Width           =   600
   End
   Begin VB.Label COMANDO 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Rew."
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   4
      Left            =   840
      TabIndex        =   20
      ToolTipText     =   "Rebobinar."
      Top             =   1680
      Width           =   600
   End
   Begin VB.Label COMANDO 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Next"
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   3
      Left            =   2955
      TabIndex        =   19
      ToolTipText     =   "Tema Siguiente."
      Top             =   1305
      Width           =   600
   End
   Begin VB.Label COMANDO 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Prev."
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   2
      Left            =   2265
      TabIndex        =   18
      ToolTipText     =   "Tema Anterior."
      Top             =   1305
      Width           =   600
   End
   Begin VB.Label COMANDO 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pause"
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   1
      Left            =   1545
      TabIndex        =   17
      ToolTipText     =   "Pausar."
      Top             =   1305
      Width           =   600
   End
   Begin VB.Label COMANDO 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Play"
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   0
      Left            =   825
      TabIndex        =   16
      ToolTipText     =   "Reproducir."
      Top             =   1305
      Width           =   600
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   4695
      Left            =   0
      Picture         =   "Main.frx":088C
      Top             =   0
      Width           =   7455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim LOCK1 As Boolean
Dim LOCK2 As Boolean
Dim n As Integer

Private Sub COMANDO_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 Select Case Index
 Case Is = 0
  If DECK1NAME <> "" Then
   Select Case PLAYERSTATUS("DECK1")
   Case Is = "stopped"
    If Form10.OPT(1).Value = 1 Then
     MIXSET(2) = 11
    Else
     Call PLAYERPLAY("DECK1", ILISTA.List(LISTA(0).ListIndex), vbNullString)
    End If
   Case Is = "playing"
    Call PLAYERSTOP("DECK1")
    Call PLAYERSETPOS("DECK1", 0)
    Call CARGAR_A
    Call PLAYERPLAY("DECK1", ILISTA.List(LISTA(0).ListIndex), vbNullString)
   Case Is = "paused"
    COMANDO(1).ForeColor = UNSELCOLOR
    Call PLAYERRESUME("DECK1")
   End Select
  End If
 Case Is = 1
  If DECK1NAME <> "" Then
   If PLAYERSTATUS("DECK1") = "playing" Then
    Call PLAYERPAUSE("DECK1")
    COMANDO(1).ForeColor = ONCOLOR
   Else
    If PLAYERSTATUS("DECK1") = "paused" Then
     Call PLAYERRESUME("DECK1")
     COMANDO(1).ForeColor = UNSELCOLOR
    End If
   End If
  End If
 Case Is = 2
  If DECK1NAME <> "" Then
   If LISTA(0).ListIndex - 1 >= 0 Then
    LISTA(0).ListIndex = LISTA(0).ListIndex - 1
    n = 0
    If PLAYERSTATUS("DECK1") = "playing" Then n = 1
    If PLAYERSTATUS("DECK1") = "paused" Then COMANDO(1).ForeColor = UNSELCOLOR
    Call CARGAR_A
    If n = 1 And DECK1NAME <> "" Then Call PLAYERPLAY("DECK1", ILISTA.List(LISTA(0).ListIndex), vbNullString)
   End If
  End If
 Case Is = 3
  If DECK1NAME <> "" Then
   If LISTA(0).ListIndex + 1 <= LISTA(0).ListCount - 1 Then
    LISTA(0).ListIndex = LISTA(0).ListIndex + 1
    n = 0
    If PLAYERSTATUS("DECK1") = "playing" Then n = 1
    If PLAYERSTATUS("DECK1") = "paused" Then COMANDO(1).ForeColor = UNSELCOLOR
    Call CARGAR_A
    If n = 1 And DECK1NAME <> "" Then Call PLAYERPLAY("DECK1", ILISTA.List(LISTA(0).ListIndex), vbNullString)
   End If
  End If
 Case Is = 4
  If DECK1NAME <> "" Then
   If Button = 1 Then
    If PLAYERGETPOS("DECK1") - 10000 >= 0 Then
     Call PLAYERSETPOS("DECK1", PLAYERGETPOS("DECK1") - 10000)
    Else
     Call PLAYERSETPOS("DECK1", 0)
    End If
   Else
    If PLAYERGETPOS("DECK1") - 1000 >= 0 Then
     Call PLAYERSETPOS("DECK1", PLAYERGETPOS("DECK1") - 1000)
    Else
     Call PLAYERSETPOS("DECK1", 0)
    End If
   End If
  End If
 Case Is = 5
  If DECK1NAME <> "" Then
   If Button = 1 Then
    If PLAYERGETPOS("DECK1") + 10000 <= PLAYERDURATION("DECK1") Then
     Call PLAYERSETPOS("DECK1", PLAYERGETPOS("DECK1") + 10000)
    Else
     Call PLAYERSETPOS("DECK1", PLAYERDURATION("DECK1") - 1000)
    End If
   Else
    If PLAYERGETPOS("DECK1") + 1000 <= PLAYERDURATION("DECK1") Then
     Call PLAYERSETPOS("DECK1", PLAYERGETPOS("DECK1") + 1000)
    Else
     Call PLAYERSETPOS("DECK1", PLAYERDURATION("DECK1"))
    End If
   End If
  End If
 Case Is = 6
  If DECK1NAME <> "" And PLAYERSTATUS("DECK1") <> "stopped" Then
   If Form10.OPT(0).Value = 1 Then
    MIXSET(2) = 12
   Else
    Call PLAYERSETPOS("DECK1", 0)
    Call PLAYERSTOP("DECK1")
    COMANDO(1).ForeColor = UNSELCOLOR
    Call CARGAR_A
   End If
  End If
 Case Is = 7
  Form2.Text1 = 0
  If LISTA(1).ListCount > 0 Then
   Form2.Drive1.Drive = SOLODRIVE(LISTA(1).List(LISTA(0).ListCount - 1))
   Form2.Dir1.path = SOLOPATH(LISTA(1).List(LISTA(0).ListCount - 1))
  Else
   If LISTB(1).ListCount > 0 Then
    Form2.Drive1.Drive = SOLODRIVE(LISTB(1).List(LISTB(0).ListCount - 1))
    Form2.Dir1.path = SOLOPATH(LISTB(1).List(LISTB(0).ListCount - 1))
   Else
    Form2.Drive1.Drive = "C:"
    Form2.Dir1.path = SOLODRIVE(App.path) + "\"
   End If
  End If
  Form2.Enabled = True
  Form2.Visible = True
  Form1.Enabled = False
  Form4.Enabled = False
  Form5.Enabled = False
 Case Is = 8
  Form2.Text1 = 1
  If LISTB(1).ListCount > 0 Then
   Form2.Drive1.Drive = SOLODRIVE(LISTB(1).List(LISTB(0).ListCount - 1))
   Form2.Dir1.path = SOLOPATH(LISTB(1).List(LISTB(0).ListCount - 1))
  Else
   If LISTA(1).ListCount > 0 Then
    Form2.Drive1.Drive = SOLODRIVE(LISTA(1).List(LISTA(0).ListCount - 1))
    Form2.Dir1.path = SOLOPATH(LISTA(1).List(LISTA(0).ListCount - 1))
   Else
    Form2.Drive1.Drive = "C:"
    Form2.Dir1.path = SOLODRIVE(App.path) + "\"
   End If
  End If
  Form2.Enabled = True
  Form2.Visible = True
  Form1.Enabled = False
  Form4.Enabled = False
  Form5.Enabled = False
 Case Is = 9
  If DECK2NAME <> "" And PLAYERSTATUS("DECK2") <> "stopped" Then
   If Form10.OPT(0).Value = 1 Then
    MIXSET(3) = 22
   Else
    Call PLAYERSETPOS("DECK2", 0)
    Call PLAYERSTOP("DECK2")
    COMANDO(14).ForeColor = UNSELCOLOR
    Call CARGAR_B
   End If
  End If
 Case Is = 11
  If DECK2NAME <> "" Then
   If Button = 1 Then
    If PLAYERGETPOS("DECK2") - 10000 >= 0 Then
     Call PLAYERSETPOS("DECK2", PLAYERGETPOS("DECK2") - 10000)
    Else
     Call PLAYERSETPOS("DECK2", 0)
    End If
   Else
    If PLAYERGETPOS("DECK2") - 1000 >= 0 Then
     Call PLAYERSETPOS("DECK2", PLAYERGETPOS("DECK2") - 1000)
    Else
     Call PLAYERSETPOS("DECK2", 0)
    End If
   End If
  End If
 Case Is = 10
  If DECK2NAME <> "" Then
   If Button = 1 Then
    If PLAYERGETPOS("DECK2") + 10000 <= PLAYERDURATION("DECK2") Then
     Call PLAYERSETPOS("DECK2", PLAYERGETPOS("DECK2") + 10000)
    Else
     Call PLAYERSETPOS("DECK2", PLAYERDURATION("DECK2") - 1000)
    End If
   Else
    If PLAYERGETPOS("DECK2") + 1000 <= PLAYERDURATION("DECK2") Then
     Call PLAYERSETPOS("DECK2", PLAYERGETPOS("DECK2") + 1000)
    Else
     Call PLAYERSETPOS("DECK2", PLAYERDURATION("DECK2"))
    End If
   End If
  End If
 Case Is = 13
  If DECK2NAME <> "" Then
   If LISTB(0).ListIndex - 1 >= 0 Then
    LISTB(0).ListIndex = LISTB(0).ListIndex - 1
    n = 0
    If PLAYERSTATUS("DECK2") = "playing" Then n = 1
    If PLAYERSTATUS("DECK2") = "paused" Then COMANDO(14).ForeColor = UNSELCOLOR
    Call CARGAR_B
    If n = 1 And DECK2NAME <> "" Then Call PLAYERPLAY("DECK2", ILISTB.List(LISTB(0).ListIndex), vbNullString)
   End If
  End If
 Case Is = 12
  If DECK2NAME <> "" Then
   If LISTB(0).ListIndex + 1 <= LISTB(0).ListCount - 1 Then
    LISTB(0).ListIndex = LISTB(0).ListIndex + 1
    n = 0
    If PLAYERSTATUS("DECK2") = "playing" Then n = 1
    If PLAYERSTATUS("DECK2") = "paused" Then COMANDO(14).ForeColor = UNSELCOLOR
    Call CARGAR_B
    If n = 1 And DECK2NAME <> "" Then Call PLAYERPLAY("DECK2", ILISTB.List(LISTB(0).ListIndex), vbNullString)
   End If
  End If
 Case Is = 15
  If DECK2NAME <> "" Then
   Select Case PLAYERSTATUS("DECK2")
   Case Is = "stopped"
    If Form10.OPT(1).Value = 1 Then
     MIXSET(3) = 21
    Else
     Call PLAYERPLAY("DECK2", ILISTB.List(LISTB(0).ListIndex), vbNullString)
    End If
   Case Is = "playing"
    Call PLAYERSTOP("DECK2")
    Call PLAYERSETPOS("DECK2", 0)
    Call CARGAR_B
    Call PLAYERPLAY("DECK2", ILISTB.List(LISTB(0).ListIndex), vbNullString)
   Case Is = "paused"
    COMANDO(14).ForeColor = UNSELCOLOR
    Call PLAYERRESUME("DECK2")
   End Select
  End If
 Case Is = 14
  If DECK2NAME <> "" Then
   If PLAYERSTATUS("DECK2") = "playing" Then
    Call PLAYERPAUSE("DECK2")
    COMANDO(14).ForeColor = ONCOLOR
   Else
    If PLAYERSTATUS("DECK2") = "paused" Then
     Call PLAYERRESUME("DECK2")
     COMANDO(14).ForeColor = UNSELCOLOR
    End If
   End If
  End If
 End Select
End Sub

Private Sub COMANDO_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 n = 0
 Do While n <= 15
  If Index = COMANDO(n).Index Then
   If COMANDO(n).ForeColor <> SELCOLOR And COMANDO(n).ForeColor <> ONCOLOR Then COMANDO(n).ForeColor = SELCOLOR
  Else
   If COMANDO(n).ForeColor <> UNSELCOLOR And COMANDO(n).ForeColor <> ONCOLOR Then COMANDO(n).ForeColor = UNSELCOLOR
  End If
  n = n + 1
 Loop
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If DECK1NAME <> "" Then
  Form6.Text2 = 0
  Form6.Enabled = True
  Form6.Visible = True
  Form1.Enabled = False
 End If
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Command1.ForeColor <> SELCOLOR Then Command1.ForeColor = SELCOLOR
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If DECK2NAME <> "" Then
  Form6.Text2 = 1
  Form6.Enabled = True
  Form6.Visible = True
  Form1.Enabled = False
 End If
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Command2.ForeColor <> SELCOLOR Then Command2.ForeColor = SELCOLOR
End Sub

Private Sub Command3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If LISTA(0).ListCount > 0 Then
  LISTA(0).Clear
  LISTA(1).Clear
  ILISTA.Clear
  FLISTA.Clear
  Call KILL_A
 End If
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Command3.ForeColor <> SELCOLOR Then Command3.ForeColor = SELCOLOR
End Sub

Private Sub Command4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 CM.Filename = ""
 CM.Filter = "Winamp Lists(*.m3u;*.pls)|*.m3u;*.pls"
 CM.ShowOpen
 If CM.CancelError = False And CM.Filename <> "" Then
  Call KILL_A
  If SOLOEXT(CM.Filename) = "pls" Then
   Call PLSLOAD(CM.Filename, LISTA(0), LISTA(1))
  Else
   Call M3ULOAD(CM.Filename, LISTA(0), LISTA(1))
  End If
  ILISTA.Clear
  FLISTA.Clear
  n = 0
  Do While n <= LISTA(0).ListCount - 1
   ILISTA.AddItem 0
   FLISTA.AddItem 0
   n = n + 1
  Loop
  LISTA(0).ListIndex = 0
  Call CARGAR_A
 End If
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Command4.ForeColor <> SELCOLOR Then Command4.ForeColor = SELCOLOR
End Sub

Private Sub Command5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If LISTB(0).ListCount > 0 Then
  LISTB(0).Clear
  LISTB(1).Clear
  ILISTB.Clear
  FLISTB.Clear
  Call KILL_B
 End If
End Sub

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Command5.ForeColor <> SELCOLOR Then Command5.ForeColor = SELCOLOR
End Sub

Private Sub Command6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 CM.Filename = ""
 CM.Filter = "Winamp Lists(*.m3u;*.pls)|*.m3u;*.pls"
 CM.ShowOpen
 If CM.CancelError = False And CM.Filename <> "" Then
  Call KILL_B
  If SOLOEXT(CM.Filename) = "pls" Then
   Call PLSLOAD(CM.Filename, LISTB(0), LISTB(1))
  Else
   Call M3ULOAD(CM.Filename, LISTB(0), LISTB(1))
  End If
  ILISTB.Clear
  FLISTB.Clear
  n = 0
  Do While n <= LISTB(0).ListCount - 1
   ILISTB.AddItem 0
   FLISTB.AddItem 0
   n = n + 1
  Loop
  LISTB(0).ListIndex = 0
  Call CARGAR_B
 End If
End Sub

Private Sub Command6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Command6.ForeColor <> SELCOLOR Then Command6.ForeColor = SELCOLOR
End Sub

Private Sub Form_Load()
 Call SetDefaultDevice("MPEGVideo", "mciqtz.drv")
 Call CloseAll
 CloseCD
 X = 0
 LOCK1 = False
 LOCK2 = False
 Form4.Visible = True
 Form4.Enabled = True
 Form5.Visible = True
 Form5.Enabled = True
 Form8.Enabled = True
 Form8.Visible = True
 Form10.Enabled = True
 Form10.Visible = True
 DECK1NAME = ""
 DECK2NAME = ""
 EFXNAME = ""
 RITMNAME = ""
 Call SetAutoRepeat(hWnd, "RITM", vbNullString, vbNullString, True)
 If Dir(App.path & "\DjBox4.ini", vbArchive) <> "" And LOADINI(INIFILE, App.path & "\DjBox4.ini") <> False Then
  Me.left = Val(INIGETVALUE(INIFILE, "LeftPosPlayer"))
  Me.top = Val(INIGETVALUE(INIFILE, "TopPosPlayer"))
  Form5.left = Val(INIGETVALUE(INIFILE, "LeftPosMixer"))
  Form5.top = Val(INIGETVALUE(INIFILE, "TopPosMixer"))
  Form4.left = Val(INIGETVALUE(INIFILE, "LeftPosFx"))
  Form4.top = Val(INIGETVALUE(INIFILE, "TopPosFx"))
  Form8.left = Val(INIGETVALUE(INIFILE, "LeftPosCD"))
  Form8.top = Val(INIGETVALUE(INIFILE, "TopPosCD"))
  Form10.left = Val(INIGETVALUE(INIFILE, "LeftPosDJ"))
  Form10.top = Val(INIGETVALUE(INIFILE, "TopPosDJ"))
  SELCOLOR = Val(INIGETVALUE(INIFILE, "ButtonSe"))
  UNSELCOLOR = Val(INIGETVALUE(INIFILE, "ButtonUn"))
  ONCOLOR = Val(INIGETVALUE(INIFILE, "ButtonOn"))
  DATAFORE = Val(INIGETVALUE(INIFILE, "DataFore"))
  DATABACK = Val(INIGETVALUE(INIFILE, "DataBack"))
  SKINFILE = Val(INIGETVALUE(INIFILE, "Skin"))
  Call Form7.SETCOLORS
  If SKINFILE > 0 Then
   Call SET_SKIN(0, SKINFILE)
   Call SET_SKIN(1, SKINFILE)
   Call SET_SKIN(2, SKINFILE)
   Call SET_SKIN(3, SKINFILE)
  End If
  Form5.WindowState = Val(INIGETVALUE(INIFILE, "StateMixer"))
  Form4.WindowState = Val(INIGETVALUE(INIFILE, "StateFx"))
  Form8.WindowState = Val(INIGETVALUE(INIFILE, "StateCD"))
  Form10.WindowState = Val(INIGETVALUE(INIFILE, "StateDJ"))
 Else
  SELCOLOR = 65280
  UNSELCOLOR = 0
  ONCOLOR = 16711680
  DATAFORE = 65280
  DATABACK = 0
 End If
 If OpenCD("d:\") = False Then
  If OpenCD("e:\") = False Then
   If OpenCD("f:\") = False Then
    If OpenCD("g:\") = False Then
     If OpenCD("h:\") = False Then MsgBox ("No Hay Una Lectora Presente En El Sistema o Está Siendo Utilizada Por Otra Aplicación.")
    End If
   End If
  End If
 End If
 If mciOpen = False Then
  Form8.Picture2.BackColor = &H80FF&
  Form8.Timer1.Enabled = False
 Else
  If MediaPresent = True Then Form8.INIT_CD
 End If
 If MsgBox("¿Desea Abrir Alguna Sesión?", vbYesNo) = vbYes Then
  CM.Filename = ""
  CM.InitDir = App.path + "\sesiones\"
  CM.Filter = "DJ4 Sesion(*.dj4)|*.dj4"
  CM.ShowOpen
  If CM.CancelError = False And CM.Filename <> "" Then
   If LoadDJ4(CM.Filename) = False Then MsgBox ("Error: No se pudo abrir la sesión.")
  End If
 End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 CURX = X
 CURY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 0 Then
  If Label10.ForeColor = &HFFFF& Then Label10.ForeColor = &HFF&
  If Label1.ForeColor = &HFFFF& Then Label1.ForeColor = &HFF&
  If Label2.ForeColor = &HFFFF& Then Label2.ForeColor = &HFF&
  If Command1.ForeColor <> UNSELCOLOR Then Command1.ForeColor = UNSELCOLOR
  If Command2.ForeColor <> UNSELCOLOR Then Command2.ForeColor = UNSELCOLOR
  If Command3.ForeColor <> UNSELCOLOR Then Command3.ForeColor = UNSELCOLOR
  If Command4.ForeColor <> UNSELCOLOR Then Command4.ForeColor = UNSELCOLOR
  If Command5.ForeColor <> UNSELCOLOR Then Command5.ForeColor = UNSELCOLOR
  If Command6.ForeColor <> UNSELCOLOR Then Command6.ForeColor = UNSELCOLOR
  n = 0
  Do While n <= 15
   If COMANDO(n).ForeColor <> UNSELCOLOR And COMANDO(n).ForeColor <> ONCOLOR Then COMANDO(n).ForeColor = UNSELCOLOR
   If n <= 3 Then
    If MARCA(n).ForeColor <> UNSELCOLOR And MARCA(n).ForeColor <> ONCOLOR Then MARCA(n).ForeColor = UNSELCOLOR
   End If
   n = n + 1
  Loop
 Else
  Me.Move Me.left + (X - CURX), Me.top + (Y - CURY)
 End If
End Sub

Private Sub Form_Resize()
 If Me.WindowState = 0 Then
  If Form4.WindowState = 1 Then Form4.WindowState = 0
  If Form5.WindowState = 1 Then Form5.WindowState = 0
  If Form8.WindowState = 1 Then Form8.WindowState = 0
  If Form10.WindowState = 1 Then Form10.WindowState = 0
 Else
  If Form4.WindowState = 0 Then Form4.WindowState = 1
  If Form5.WindowState = 0 Then Form5.WindowState = 1
  If Form8.WindowState = 0 Then Form8.WindowState = 1
  If Form10.WindowState = 0 Then Form10.WindowState = 1
 End If
End Sub

Private Sub Label1_Click()
 Me.WindowState = 1
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Label1.ForeColor <> &HFFFF& Then Label1.ForeColor = &HFFFF&
 If Label10.ForeColor = &HFFFF& Then Label10.ForeColor = &HFF&
 If Label2.ForeColor = &HFFFF& Then Label2.ForeColor = &HFF&
End Sub

Private Sub Label10_Click()
 If Dir(App.path & "\DjBox4.ini", vbArchive) <> "" Then
  Call INISETVALUE(INIFILE, "StateMixer", Form5.WindowState)
  Call INISETVALUE(INIFILE, "StateFx", Form4.WindowState)
  Call INISETVALUE(INIFILE, "StateCD", Form8.WindowState)
  Call INISETVALUE(INIFILE, "StateDJ", Form10.WindowState)
  Call INISETVALUE(INIFILE, "LeftPosPlayer", Form1.left)
  Call INISETVALUE(INIFILE, "TopPosPlayer", Form1.top)
  Call INISETVALUE(INIFILE, "LeftPosMixer", Form5.left)
  Call INISETVALUE(INIFILE, "TopPosMixer", Form5.top)
  Call INISETVALUE(INIFILE, "LeftPosFx", Form4.left)
  Call INISETVALUE(INIFILE, "TopPosFx", Form4.top)
  Call INISETVALUE(INIFILE, "LeftPosCD", Form8.left)
  Call INISETVALUE(INIFILE, "TopPosCD", Form8.top)
  Call INISETVALUE(INIFILE, "LeftPosDJ", Form10.left)
  Call INISETVALUE(INIFILE, "TopPosDJ", Form10.top)
  Call INISETVALUE(INIFILE, "ButtonSe", str(SELCOLOR))
  Call INISETVALUE(INIFILE, "ButtonUn", str(UNSELCOLOR))
  Call INISETVALUE(INIFILE, "ButtonOn", str(ONCOLOR))
  Call INISETVALUE(INIFILE, "DataFore", str(DATAFORE))
  Call INISETVALUE(INIFILE, "DataBack", str(DATABACK))
  Call INISETVALUE(INIFILE, "Skin", str(SKINFILE))
  Call SAVEINI(INIFILE, App.path & "\DjBox4.ini")
 End If
 If mciOpen = True Then
  If IsStopped = False Then StopCD
  CloseCD
 End If
 Call CloseAll
 If LISTA(0).ListCount > 0 Or LISTB(0).ListCount > 0 Then
  If MsgBox("¿Desea Guardar La Sesión?", vbYesNo) = vbYes Then
   CM.Filename = ""
   CM.InitDir = App.path + "\sesiones\"
   CM.Filter = "DJ4 Sesion(*.dj4)|*.dj4"
   CM.ShowSave
   If CM.CancelError = False And CM.Filename <> "" Then
    If SaveDJ4(CM.Filename) = False Then MsgBox ("Error: No se pudo guardar la sesión.")
   End If
  End If
 End If
 End
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Label10.ForeColor <> &HFFFF& Then Label10.ForeColor = &HFFFF&
 If Label1.ForeColor = &HFFFF& Then Label1.ForeColor = &HFF&
 If Label2.ForeColor = &HFFFF& Then Label2.ForeColor = &HFF&
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Form7.Enabled = True
 Form7.Visible = True
 Form1.Enabled = False
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Label2.ForeColor <> &HFFFF& Then Label2.ForeColor = &HFFFF&
 If Label1.ForeColor = &HFFFF& Then Label1.ForeColor = &HFF&
 If Label10.ForeColor = &HFFFF& Then Label10.ForeColor = &HFF&
End Sub

Private Sub LISTA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 If LISTA(0).ListCount > 0 Then
  If Button = 1 Then
   n = 0
   If PLAYERSTATUS("DECK1") = "playing" Then n = 1
   If PLAYERSTATUS("DECK1") = "paused" Then COMANDO(1).ForeColor = UNSELCOLOR
   Call CARGAR_A
   If n = 1 And DECK1NAME <> "" Then Call PLAYERPLAY("DECK1", ILISTA.List(LISTA(0).ListIndex), vbNullString)
  Else
   If UCase(SOLOEXT(LISTA(1).List(LISTA(0).ListIndex))) = "MP3" Then
    Call getMP3Info(LISTA(1).List(LISTA(0).ListIndex), MP3INF)
    MsgBox ("Bitrate: " & MP3INF.BITRATE & Chr(10) & "Channels: " & MP3INF.CHANNELS & Chr(10) & "Copyright: " & MP3INF.COPYRIGHT & Chr(10) & "CRC: " & MP3INF.CRC & Chr(10) & "Emphasis: " & MP3INF.EMPHASIS & Chr(10) & "Freq: " & MP3INF.FREQ & Chr(10) & "Layer: " & MP3INF.LAYER & Chr(10) & "Lenght: " & MP3INF.LENGTH & Chr(10) & "Mpeg: " & MP3INF.MPEG & Chr(10) & "Original: " & MP3INF.ORIGINAL & Chr(10) & "Size: " & MP3INF.SIZE)
   Else
    MsgBox (WAVEINFO(LISTA(1).List(LISTA(0).ListIndex)))
   End If
  End If
 End If
End Sub

Private Sub LISTB_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 If LISTB(0).ListCount > 0 Then
  If Button = 1 Then
   n = 0
   If PLAYERSTATUS("DECK2") = "playing" Then n = 1
   If PLAYERSTATUS("DECK2") = "paused" Then COMANDO(14).ForeColor = UNSELCOLOR
   Call CARGAR_B
   If n = 1 And DECK2NAME <> "" Then Call PLAYERPLAY("DECK2", ILISTB.List(LISTB(0).ListIndex), vbNullString)
  Else
   If UCase(SOLOEXT(LISTB(1).List(LISTB(0).ListIndex))) = "MP3" Then
    Call getMP3Info(LISTB(1).List(LISTB(0).ListIndex), MP3INF)
    MsgBox ("Bitrate: " & MP3INF.BITRATE & Chr(10) & "Channels: " & MP3INF.CHANNELS & Chr(10) & "Copyright: " & MP3INF.COPYRIGHT & Chr(10) & "CRC: " & MP3INF.CRC & Chr(10) & "Emphasis: " & MP3INF.EMPHASIS & Chr(10) & "Freq: " & MP3INF.FREQ & Chr(10) & "Layer: " & MP3INF.LAYER & Chr(10) & "Lenght: " & MP3INF.LENGTH & Chr(10) & "Mpeg: " & MP3INF.MPEG & Chr(10) & "Original: " & MP3INF.ORIGINAL & Chr(10) & "Size: " & MP3INF.SIZE)
   Else
    MsgBox (WAVEINFO(LISTB(1).List(LISTB(0).ListIndex)))
   End If
  End If
 End If
End Sub

Private Sub MARCA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 Select Case Index
 Case Is = 0
  If DECK1NAME <> "" Then
   If MARCA(0).ForeColor = ONCOLOR Then
    ILISTA.List(LISTA(0).ListIndex) = 0
    MARCA(0).ForeColor = UNSELCOLOR
   Else
    If PLAYERGETPOS("DECK1") > 0 Then
     ILISTA.List(LISTA(0).ListIndex) = PLAYERGETPOS("DECK1")
     MARCA(0).ForeColor = ONCOLOR
    End If
   End If
  End If
 Case Is = 1
  If DECK1NAME <> "" Then
   If MARCA(1).ForeColor = ONCOLOR Then
    FLISTA.List(LISTA(0).ListIndex) = 0
    MARCA(1).ForeColor = UNSELCOLOR
   Else
    If PLAYERGETPOS("DECK1") < PLAYERDURATION("DECK1") Then
     FLISTA.List(LISTA(0).ListIndex) = PLAYERGETPOS("DECK1")
     MARCA(1).ForeColor = ONCOLOR
    End If
   End If
  End If
 Case Is = 2
  If DECK2NAME <> "" Then
   If MARCA(2).ForeColor = ONCOLOR Then
    ILISTB.List(LISTB(0).ListIndex) = 0
    MARCA(2).ForeColor = UNSELCOLOR
   Else
    If PLAYERGETPOS("DECK2") > 0 Then
     ILISTB.List(LISTB(0).ListIndex) = PLAYERGETPOS("DECK2")
     MARCA(2).ForeColor = ONCOLOR
    End If
   End If
  End If
 Case Is = 3
  If DECK2NAME <> "" Then
   If MARCA(3).ForeColor = ONCOLOR Then
    FLISTB.List(LISTB(0).ListIndex) = 0
    MARCA(3).ForeColor = UNSELCOLOR
   Else
    If PLAYERGETPOS("DECK2") < PLAYERDURATION("DECK2") Then
     FLISTB.List(LISTB(0).ListIndex) = PLAYERGETPOS("DECK2")
     MARCA(3).ForeColor = ONCOLOR
    End If
   End If
  End If
 End Select
End Sub

Private Sub MARCA_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 n = 0
 Do While n <= 3
  If Index = MARCA(n).Index Then
   If MARCA(n).ForeColor <> SELCOLOR And MARCA(n).ForeColor <> ONCOLOR Then MARCA(n).ForeColor = SELCOLOR
  Else
   If MARCA(n).ForeColor <> UNSELCOLOR And MARCA(n).ForeColor <> ONCOLOR Then MARCA(n).ForeColor = UNSELCOLOR
  End If
  n = n + 1
 Loop
End Sub

Private Sub TIMEA_Change()
 If LOCK1 = True And Int(PLAYERGETPOS("DECK1") / 1000) <> TIMEA.Value Then Call PLAYERSETPOS("DECK1", TIMEA.Value * 1000)
 If LOCK1 = True Then LOCK1 = False
End Sub

Private Sub TIMEA_Scroll()
 If LOCK1 = False Then LOCK1 = True
 CURPOSA.Caption = MINSEG(TIMEA.Value)
End Sub

Private Sub TIMEB_Change()
 If LOCK2 = True And Int(PLAYERGETPOS("DECK2") / 1000) <> TIMEB.Value Then Call PLAYERSETPOS("DECK2", TIMEB.Value * 1000)
 If LOCK2 = True Then LOCK2 = False
End Sub

Private Sub TIMEB_Scroll()
 If LOCK2 = False Then LOCK2 = True
 CURPOSB.Caption = MINSEG(TIMEB.Value)
End Sub

Private Sub Timer1_Timer()
 If Me.WindowState = 0 Then
  If DECK1NAME <> "" Then
   If LOCK1 = False Then
    If CURPOSA <> "< " + MINSEG(PLAYERGETPOS("DECK1") / 1000) + " >" Then CURPOSA = "< " + MINSEG(PLAYERGETPOS("DECK1") / 1000) + " >"
    If TIMEA.Value <> Int(PLAYERGETPOS("DECK1") / 1000) Then TIMEA.Value = Int(PLAYERGETPOS("DECK1") / 1000)
   End If
   If ILISTA.List(LISTA(0).ListIndex) <> 0 Then
    If INICIOA <> MINSEG(ILISTA.List(LISTA(0).ListIndex) / 1000) Then INICIOA = MINSEG(ILISTA.List(LISTA(0).ListIndex) / 1000)
   Else
    INICIOA = "00:00"
   End If
   If FLISTA.List(LISTA(0).ListIndex) <> 0 Then
    If FINA <> MINSEG(FLISTA.List(LISTA(0).ListIndex) / 1000) Then FINA = MINSEG(FLISTA.List(LISTA(0).ListIndex) / 1000)
   Else
    If FINA <> MINSEG(PLAYERDURATION("DECK1") / 1000) Then FINA = MINSEG(PLAYERDURATION("DECK1") / 1000)
   End If
   Select Case PLAYERSTATUS("DECK1")
   Case Is = "stopped"
    If STATEBMP(0).Picture <> STATECLIP.GraphicCell(1) Then STATEBMP(0).Picture = STATECLIP.GraphicCell(1)
   Case Is = "playing"
    If STATEBMP(0).Picture <> STATECLIP.GraphicCell(0) Then STATEBMP(0).Picture = STATECLIP.GraphicCell(0)
   Case Is = "paused"
    If STATEBMP(0).Picture <> STATECLIP.GraphicCell(2) Then STATEBMP(0).Picture = STATECLIP.GraphicCell(2)
   End Select
  End If
  If DECK2NAME <> "" Then
   If LOCK2 = False Then
    If CURPOSB <> "< " + MINSEG(PLAYERGETPOS("DECK2") / 1000) + " >" Then CURPOSB = "< " + MINSEG(PLAYERGETPOS("DECK2") / 1000) + " >"
    If TIMEB.Value <> Int(PLAYERGETPOS("DECK2") / 1000) Then TIMEB.Value = Int(PLAYERGETPOS("DECK2") / 1000)
   End If
   If Val(ILISTB.List(LISTB(0).ListIndex)) <> 0 Then
    If INICIOB <> MINSEG(ILISTB.List(LISTB(0).ListIndex) / 1000) Then INICIOB = MINSEG(ILISTB.List(LISTB(0).ListIndex) / 1000)
   Else
    INICIOB = "00:00"
   End If
   If Val(FLISTB.List(LISTB(0).ListIndex)) <> 0 Then
    If FINB <> MINSEG(FLISTB.List(LISTB(0).ListIndex) / 1000) Then FINB = MINSEG(FLISTB.List(LISTB(0).ListIndex) / 1000)
   Else
    If FINB <> MINSEG(PLAYERDURATION("DECK2") / 1000) Then FINB = MINSEG(PLAYERDURATION("DECK2") / 1000)
   End If
   Select Case PLAYERSTATUS("DECK2")
   Case Is = "stopped"
    If STATEBMP(1).Picture <> STATECLIP.GraphicCell(1) Then STATEBMP(1).Picture = STATECLIP.GraphicCell(1)
   Case Is = "playing"
    If STATEBMP(1).Picture <> STATECLIP.GraphicCell(0) Then STATEBMP(1).Picture = STATECLIP.GraphicCell(0)
   Case Is = "paused"
    If STATEBMP(1).Picture <> STATECLIP.GraphicCell(2) Then STATEBMP(1).Picture = STATECLIP.GraphicCell(2)
   End Select
  End If
 End If
End Sub

Public Sub CARGAR_A()
 If DECK1NAME <> LISTA(1).List(LISTA(0).ListIndex) Then
  Call PLAYERCLOSE("DECK1")
  If Dir(LISTA(1).List(LISTA(0).ListIndex), vbArchive) <> "" Then
   Call PLAYEROPEN(hWnd, "DECK1", LISTA(1).List(LISTA(0).ListIndex), "MPEGVideo")
  Else
   MsgBox ("Error04: El Archivo No Existe")
  End If
 End If
 If DECK1NAME <> "" Then
  Call Form5.VOL1
  Call VELOCC1
  If FLISTA.List(LISTA(0).ListIndex) > 0 And MARCA(1).ForeColor <> &HFF0000 Then MARCA(1).ForeColor = &HFF0000
  If FLISTA.List(LISTA(0).ListIndex) = 0 And MARCA(1).ForeColor <> &H0& Then MARCA(1).ForeColor = &H0&
  DURA = "- | " + MINSEG(PLAYERDURATION("DECK1") / 1000) + " |"
  If Int(PLAYERDURATION("DECK1") / 1000) > 0 Then
   TIMEA.Max = Int(PLAYERDURATION("DECK1") / 1000)
  Else
   TIMEA.Max = 1
   TIMEA.Value = 0
  End If
  If ILISTA.List(LISTA(0).ListIndex) > 0 Then
   If MARCA(0).ForeColor <> ONCOLOR Then MARCA(0).ForeColor = ONCOLOR
  Else
   If MARCA(0).ForeColor <> UNSELCOLOR Then MARCA(0).ForeColor = UNSELCOLOR
  End If
 Else
  MsgBox ("Error02: El archivo de audio no pudo ser cargado debido a que posiblemente esté dañado o hay otra aplicación utilizando el dispositivo MCI.")
 End If
 NUMA = PISTA(LISTA(0).ListIndex + 1, 4)
 TITLEA.Caption = Chr(34) + " " + LISTA(0).List(LISTA(0).ListIndex) + " " + Chr(34)
End Sub

Public Sub CARGAR_B()
 If DECK2NAME <> LISTB(1).List(LISTB(0).ListIndex) Then
  Call PLAYERCLOSE("DECK2")
  Call PLAYEROPEN(hWnd, "DECK2", LISTB(1).List(LISTB(0).ListIndex), "MPEGVideo")
 End If
 If DECK2NAME <> "" Then
  Call Form5.VOL2
  Call VELOCC2
  If Val(FLISTB.List(LISTB(0).ListIndex)) > 0 And MARCA(3).ForeColor <> &HFF0000 Then MARCA(3).ForeColor = &HFF0000
  If Val(FLISTB.List(LISTB(0).ListIndex)) = 0 And MARCA(3).ForeColor <> &H0& Then MARCA(3).ForeColor = &H0&
  DURB = "- | " + MINSEG(PLAYERDURATION("DECK2") / 1000) + " |"
  If Int(PLAYERDURATION("DECK2")) > 0 Then TIMEB.Max = Int(PLAYERDURATION("DECK2") / 1000)
  If Val(ILISTB.List(LISTB(0).ListIndex)) > 0 Then
   If MARCA(2).ForeColor <> ONCOLOR Then MARCA(2).ForeColor = ONCOLOR
  Else
   If MARCA(2).ForeColor <> UNSELCOLOR Then MARCA(2).ForeColor = UNSELCOLOR
  End If
 Else
  MsgBox ("Error02: El archivo de audio no pudo ser cargado debido a que posiblemente esté dañado o hay otra aplicación utilizando el dispositivo MCI.")
 End If
 NUMB = PISTA(LISTB(0).ListIndex + 1, 4)
 TITLEB.Caption = Chr(34) + " " + LISTB(0).List(LISTB(0).ListIndex) + " " + Chr(34)
End Sub

Private Sub TIMERFX_Timer()
 Text1 = Form10.LOOPBUT.ToolTipText
 Text2 = PLAYERGETPOS("DECK1")
 If LISTA(0).ListCount > 0 And LISTA(0).ListIndex <= 0 Then LISTA(0).ListIndex = 0
 If LISTB(0).ListCount > 0 And LISTB(0).ListIndex <= 0 Then LISTB(0).ListIndex = 0
 Select Case MIXSET(1)
 Case Is = 11
  Call MIX1
 Case Is = 21
  Call MIX2
 Case Is = 31
  Call MIX3
 Case Is = 12
  Call MIX1
 Case Is = 22
  Call MIX2
 Case Is = 32
  Call MIX3
 End Select
 If MIXSET(2) = 11 Then FADEIN1
 If MIXSET(3) = 21 Then FADEIN2
 If MIXSET(2) = 12 Then FADEOUT1
 If MIXSET(3) = 22 Then FADEOUT2
 If MIXSET(5) > 0 Then CROSSAB
 If Form10.OPT(2).Value = 1 And MIXSET(4) = 0 And PLAYERSTATUS("EFX") = "playing" And PLAYERGETPOS("EFX") >= PLAYERDURATION("EFX") - (PLAYERDURATION("EFX") / 3) And AreMultimediaAtEnd("EFX", 0) = False Then MIXSET(4) = 1
 If MIXSET(4) = 1 Then Call FADEOUT3
End Sub

Private Sub VEL_Change(Index As Integer)
 If Index = 0 Then
  Call VELOCC1
 Else
  Call VELOCC2
 End If
End Sub

Private Sub VEL_Scroll(Index As Integer)
 If Index = 0 Then
  Select Case VEL(0)
  Case Is < 20
   ACTION(0).Caption = "-" + Trim(str(20 - VEL(0)))
  Case Is = 20
   ACTION(0).Caption = "0"
  Case Is > 20
   ACTION(0).Caption = "+" + Trim(str(VEL(0) - 20))
  End Select
 Else
  Select Case VEL(1)
  Case Is < 20
   ACTION(1).Caption = "-" + Trim(str(20 - VEL(1)))
  Case Is = 20
   ACTION(1).Caption = "0"
  Case Is > 20
   ACTION(1).Caption = "+" + Trim(str(VEL(1) - 20))
  End Select
 End If
End Sub

Private Sub KILL_B()
 If DECK2NAME <> "" Then
  If PLAYERSTATUS("DECK2") = "playing" Or PLAYERSTATUS("DECK2") = "paused" Then
   Call PLAYERSETPOS("DECK2", 0)
   Call PLAYERSTOP("DECK2")
   If COMANDO(14).ForeColor <> UNSELCOLOR Then COMANDO(14).ForeColor = UNSELCOLOR
  End If
  Call PLAYERCLOSE("DECK2")
  If TIMEB.Value > 0 Then TIMEB.Value = 0
  If NUMB <> "0000" Then NUMB = "0000"
  If TITLEB <> Chr(34) + " Vacío " + Chr(34) Then TITLEB = Chr(34) + " Vacío " + Chr(34)
  If CURPOSB <> "< 00:00 >" Then CURPOSB = "< 00:00 >"
  If DURB <> "- | 00:00 |" Then DURB = "- | 00:00 |"
  If INICIOB <> "00:00" Then INICIOB = "00:00"
  If FINB <> "00:00" Then FINB = "00:00"
  If MARCA(2).ForeColor <> &H0& Then MARCA(2).ForeColor = &H0&
  If MARCA(3).ForeColor <> &H0& Then MARCA(3).ForeColor = &H0&
 End If
End Sub

Private Sub KILL_A()
 If DECK1NAME <> "" Then
  If PLAYERSTATUS("DECK1") = "playing" Or PLAYERSTATUS("DECK1") = "paused" Then
   Call PLAYERSETPOS("DECK1", 0)
   Call PLAYERSTOP("DECK1")
   If COMANDO(1).ForeColor <> UNSELCOLOR Then COMANDO(1).ForeColor = UNSELCOLOR
  End If
  Call PLAYERCLOSE("DECK1")
  If TIMEA.Value > 0 Then TIMEA.Value = 0
  If NUMA <> "0000" Then NUMA = "0000"
  If TITLEA <> Chr(34) + " Vacío " + Chr(34) Then TITLEA = Chr(34) + " Vacío " + Chr(34)
  If CURPOSA <> "< 00:00 >" Then CURPOSA = "< 00:00 >"
  If DURA <> "- | 00:00 |" Then DURA = "- | 00:00 |"
  If INICIOA <> "00:00" Then INICIOA = "00:00"
  If FINA <> "00:00" Then FINA = "00:00"
  If MARCA(0).ForeColor <> &H0& Then MARCA(0).ForeColor = &H0&
  If MARCA(1).ForeColor <> &H0& Then MARCA(1).ForeColor = &H0&
 End If
End Sub

Public Sub SWITCH()
 If DECK1NAME <> "" And DECK2NAME <> "" Then
  If PLAYERSTATUS("DECK1") <> "stopped" And PLAYERSTATUS("DECK2") <> "playing" Then
   If Form10.Command1(0).ForeColor <> ONCOLOR Then Form10.Command1(0).ForeColor = ONCOLOR
   Call PLAYERPLAY("DECK2", ILISTB.List(LISTB(0).ListIndex), vbNullString)
   Call PLAYERSETPOS("DECK1", 0)
   Call PLAYERSTOP("DECK1")
   If Form10.Command1(0).ForeColor <> UNSELCOLOR Then Form10.Command1(0).ForeColor = UNSELCOLOR
  Else
   If PLAYERSTATUS("DECK2") <> "stopped" And PLAYERSTATUS("DECK1") <> "playing" Then
    If Form10.Command1(0).ForeColor <> ONCOLOR Then Form10.Command1(0).ForeColor = ONCOLOR
    Call PLAYERPLAY("DECK1", ILISTA.List(LISTA(0).ListIndex), vbNullString)
    Call PLAYERSETPOS("DECK2", 0)
    Call PLAYERSTOP("DECK2")
    If Form10.Command1(0).ForeColor <> UNSELCOLOR Then Form10.Command1(0).ForeColor = UNSELCOLOR
   End If
  End If
 End If
End Sub

Private Sub MIX1()
 Dim VELOCC As Integer
 VELOCC = (11 - Form10.MIXVEL.Value) * 2
 If Form10.Command1(1).ForeColor <> ONCOLOR Then Form10.Command1(1).ForeColor = ONCOLOR
 If MIXSET(1) = 11 Then
  If PLAYERSTATUS("DECK2") <> "playing" Then Call PLAYERPLAY("DECK2", ILISTB.List(LISTB(0).ListIndex), vbNullString)
  If PLAYERGETVOLUME("DECK1", "all") - VELOCC >= 0 Then
   Call PLAYERSETVOLUME("DECK1", "all", PLAYERGETVOLUME("DECK1", "all") - VELOCC)
   If Form10.LOOPBUT.ForeColor = ONCOLOR Then
    If Form10.LOOPBUT.ToolTipText = 0 Then
     Form10.LOOPBUT.ToolTipText = PLAYERGETPOS("DECK1")
     Call PLAYERSETPOS("DECK1", PLAYERGETPOS("DECK1") - 5000)
    Else
     If Trim(Form10.LOOPBUT.ToolTipText) <= Trim(PLAYERGETPOS("DECK1")) Then Call PLAYERSETPOS("DECK1", PLAYERGETPOS("DECK1") - 5000)
    End If
   End If
  Else
   Call PLAYERSETPOS("DECK1", 0)
   Call PLAYERSTOP("DECK1")
   MIXSET(1) = 0
   Call Form5.VOL1
   If Form10.Command1(1).ForeColor <> UNSELCOLOR Then Form10.Command1(1).ForeColor = UNSELCOLOR
  End If
 Else
  If PLAYERSTATUS("DECK1") <> "playing" Then Call PLAYERPLAY("DECK1", ILISTA.List(LISTA(0).ListIndex), vbNullString)
  If PLAYERGETVOLUME("DECK2", "all") - VELOCC >= 0 Then
   Call PLAYERSETVOLUME("DECK2", "all", PLAYERGETVOLUME("DECK2", "all") - VELOCC)
   If Form10.LOOPBUT.ForeColor = ONCOLOR Then
    If Form10.LOOPBUT.ToolTipText = 0 Then
     Form10.LOOPBUT.ToolTipText = PLAYERGETPOS("DECK2")
     Call PLAYERSETPOS("DECK2", PLAYERGETPOS("DECK2") - 10000)
    Else
     If Form10.LOOPBUT.ToolTipText <= PLAYERGETPOS("DECK1") Then Call PLAYERSETPOS("DECK2", PLAYERGETPOS("DECK2") - 10000)
    End If
   End If
  Else
   Call PLAYERSETPOS("DECK2", 0)
   Call PLAYERSTOP("DECK2")
   MIXSET(1) = 0
   Call Form5.VOL2
   If Form10.Command1(1).ForeColor <> UNSELCOLOR Then Form10.Command1(1).ForeColor = UNSELCOLOR
  End If
 End If
End Sub

Private Sub MIX2()
 Dim VELOCC As Integer
 VELOCC = (11 - Form10.MIXVEL.Value) * 2
 If Form10.Command1(2).ForeColor <> ONCOLOR Then Form10.Command1(2).ForeColor = ONCOLOR
 If MIXSET(1) = 21 Then
  If PLAYERGETVOLUME("DECK1", "all") - VELOCC >= 0 Then
   Call PLAYERSETVOLUME("DECK1", "all", PLAYERGETVOLUME("DECK1", "all") - VELOCC)
  Else
   Call PLAYERSETPOS("DECK1", 0)
   Call PLAYERSTOP("DECK1")
   MIXSET(1) = 0
   Call Form5.VOL1
   If Form10.Command1(2).ForeColor <> UNSELCOLOR Then Form10.Command1(2).ForeColor = UNSELCOLOR
  End If
  If PLAYERGETVOLUME("DECK1", "all") <= 20 And PLAYERSTATUS("DECK2") <> "playing" Then Call PLAYERPLAY("DECK2", ILISTB.List(LISTB(0).ListIndex), vbNullString)
 Else
  If PLAYERGETVOLUME("DECK2", "all") - VELOCC >= 0 Then
   Call PLAYERSETVOLUME("DECK2", "all", PLAYERGETVOLUME("DECK2", "all") - VELOCC)
  Else
   Call PLAYERSETPOS("DECK2", 0)
   Call PLAYERSTOP("DECK2")
   MIXSET(1) = 0
   Call Form5.VOL2
   If Form10.Command1(2).ForeColor <> UNSELCOLOR Then Form10.Command1(2).ForeColor = UNSELCOLOR
  End If
  If PLAYERGETVOLUME("DECK2", "all") <= 20 And PLAYERSTATUS("DECK1") <> "playing" Then Call PLAYERPLAY("DECK1", ILISTA.List(LISTA(0).ListIndex), vbNullString)
 End If
End Sub

Private Sub MIX3()
 Dim capo As Long
 Dim VELOCC As Integer
 If Form10.Command1(3).ForeColor <> ONCOLOR Then Form10.Command1(3).ForeColor = ONCOLOR
 VELOCC = (11 - Form10.MIXVEL.Value) * 2
 If MIXSET(1) = 31 Then
  If Form5.CROSSFADE.Value >= 100 Then
   capo = Int(((100 - Form5.VOL(0).Value) * (100 - Form5.VOL(2).Value)) / 100)
  Else
   capo = Int((Form5.CROSSFADE.Value * Int(((100 - Form5.VOL(0).Value) * (100 - Form5.VOL(2).Value)) / 100)) / 100)
  End If
  If PLAYERSTATUS("DECK2") <> "playing" Then
   Call PLAYERSETVOLUME("DECK2", "all", 0)
   Call PLAYERPLAY("DECK2", ILISTB.List(LISTB(0).ListIndex), vbNullString)
  End If
  If PLAYERGETVOLUME("DECK1", "all") - VELOCC >= 0 Then
   Call PLAYERSETVOLUME("DECK1", "all", PLAYERGETVOLUME("DECK1", "all") - VELOCC)
  Else
   Call PLAYERSETVOLUME("DECK1", "all", 0)
  End If
  If PLAYERGETVOLUME("DECK2", "all") + VELOCC <= capo Then
   Call PLAYERSETVOLUME("DECK2", "all", PLAYERGETVOLUME("DECK2", "all") + VELOCC)
  Else
   Call PLAYERSETVOLUME("DECK2", "all", capo)
  End If
  If PLAYERGETVOLUME("DECK2", "all") = capo And PLAYERGETVOLUME("DECK1", "all") = 0 Then
   Call PLAYERSETPOS("DECK1", 0)
   Call PLAYERSTOP("DECK1")
   MIXSET(1) = 0
   Call Form5.VOL1
   If Form10.Command1(3).ForeColor <> UNSELCOLOR Then Form10.Command1(3).ForeColor = UNSELCOLOR
  End If
 Else
  If Form5.CROSSFADE.Value >= 100 Then
   capo = Int(((100 - Form5.VOL(0).Value) * (100 - Form5.VOL(1).Value)) / 100)
  Else
   capo = Int((Form5.CROSSFADE.Value * Int(((100 - Form5.VOL(0).Value) * (100 - Form5.VOL(1).Value)) / 100)) / 100)
  End If
  If PLAYERSTATUS("DECK1") <> "playing" Then
   Call PLAYERSETVOLUME("DECK1", "all", 0)
   Call PLAYERPLAY("DECK1", ILISTA.List(LISTA(0).ListIndex), vbNullString)
  End If
  If PLAYERGETVOLUME("DECK2", "all") - VELOCC >= 0 Then
   Call PLAYERSETVOLUME("DECK2", "all", PLAYERGETVOLUME("DECK2", "all") - VELOCC)
  Else
   Call PLAYERSETVOLUME("DECK2", "all", 0)
  End If
  If PLAYERGETVOLUME("DECK1", "all") + VELOCC <= capo Then
   Call PLAYERSETVOLUME("DECK1", "all", PLAYERGETVOLUME("DECK1", "all") + VELOCC)
  Else
   Call PLAYERSETVOLUME("DECK1", "all", capo)
  End If
  If PLAYERGETVOLUME("DECK1", "all") = capo And PLAYERGETVOLUME("DECK2", "all") = 0 Then
   Call PLAYERSETPOS("DECK2", 0)
   Call PLAYERSTOP("DECK2")
   MIXSET(1) = 0
   Call Form5.VOL2
   If Form10.Command1(3).ForeColor <> UNSELCOLOR Then Form10.Command1(3).ForeColor = UNSELCOLOR
  End If
 End If
End Sub

Private Sub FADEOUT1()
 If PLAYERGETVOLUME("DECK1", "all") > 0 Then
  If (PLAYERGETVOLUME("DECK1", "all") - 10) < 0 Then
   Call PLAYERSETVOLUME("DECK1", "all", 0)
   Call PLAYERSETPOS("DECK1", 0)
   Call PLAYERSTOP("DECK1")
   MIXSET(2) = 0
   Call Form5.VOL1
  Else
   Call PLAYERSETVOLUME("DECK1", "all", PLAYERGETVOLUME("DECK1", "all") - 10)
  End If
 End If
End Sub

Private Sub FADEOUT2()
 If PLAYERGETVOLUME("DECK2", "all") > 0 Then
  If (PLAYERGETVOLUME("DECK2", "all") - 10) < 0 Then
   Call PLAYERSETVOLUME("DECK2", "all", 0)
   Call PLAYERSETPOS("DECK2", 0)
   Call PLAYERSTOP("DECK2")
   MIXSET(3) = 0
   Call Form5.VOL2
  Else
   Call PLAYERSETVOLUME("DECK2", "all", PLAYERGETVOLUME("DECK2", "all") - 10)
  End If
 End If
End Sub

Private Sub FADEIN1()
 Dim capo As Long
 If Form5.CROSSFADE.Value >= 100 Then
  capo = Int(((100 - Form5.VOL(0).Value) * (100 - Form5.VOL(1).Value)) / 100)
 Else
  capo = Int((Form5.CROSSFADE.Value * Int(((100 - Form5.VOL(0).Value) * (100 - Form5.VOL(1).Value)) / 100)) / 100)
 End If
 If PLAYERSTATUS("DECK1") <> "playing" Then
  Call PLAYERSETVOLUME("DECK1", "all", 0)
  Call PLAYERPLAY("DECK1", ILISTA.List(LISTA(0).ListIndex), vbNullString)
 End If
 If PLAYERGETVOLUME("DECK1", "all") + 10 < capo Then
  Call PLAYERSETVOLUME("DECK1", "all", PLAYERGETVOLUME("DECK1", "all") + 10)
 Else
  Call PLAYERSETVOLUME("DECK1", "all", capo)
  MIXSET(2) = 0
 End If
End Sub

Private Sub FADEIN2()
 Dim capo As Long
 If Form5.CROSSFADE.Value >= 100 Then
  capo = Int(((100 - Form5.VOL(0).Value) * (100 - Form5.VOL(2).Value)) / 100)
 Else
  capo = Int((Form5.CROSSFADE.Value * Int(((100 - Form5.VOL(0).Value) * (100 - Form5.VOL(2).Value)) / 100)) / 100)
 End If
 If PLAYERSTATUS("DECK2") <> "playing" Then
  Call PLAYERSETVOLUME("DECK2", "all", 0)
  Call PLAYERPLAY("DECK2", ILISTB.List(LISTB(0).ListIndex), vbNullString)
 End If
 If PLAYERGETVOLUME("DECK2", "all") + 10 < capo Then
  Call PLAYERSETVOLUME("DECK2", "all", PLAYERGETVOLUME("DECK2", "all") + 10)
 Else
  Call PLAYERSETVOLUME("DECK2", "all", capo)
  MIXSET(3) = 0
 End If
End Sub

Private Sub FADEOUT3()
 If PLAYERGETVOLUME("EFX", "all") > 0 Then
  If (PLAYERGETVOLUME("EFX", "all") - 10) < 0 Then
   Call PLAYERSETVOLUME("EFX", "all", 0)
  Else
   Call PLAYERSETVOLUME("EFX", "all", PLAYERGETVOLUME("EFX", "all") - 10)
  End If
  If AreMultimediaAtEnd("EFX", 0) = True Then
   MIXSET(4) = 0
   Call Form5.VOL3
  End If
 End If
End Sub

Private Sub CROSSAB()
 If MIXSET(5) = 1 Then
  If Form5.CROSSFADE.Value - 10 >= 0 Then
   Form5.CROSSFADE.Value = Form5.CROSSFADE.Value - 10
  Else
   Form5.CROSSFADE.Value = 0
   MIXSET(5) = 0
  End If
 Else
  If Form5.CROSSFADE.Value + 10 <= 200 Then
   Form5.CROSSFADE.Value = Form5.CROSSFADE.Value + 10
  Else
   Form5.CROSSFADE.Value = 200
   MIXSET(5) = 0
  End If
 End If
 Call Form5.CROSSFADE_Scroll
End Sub

Private Sub VELOCC1()
 Select Case VEL(0)
 Case Is < 20
  Call PLAYERSETRATE("DECK1", 100 - (20 - VEL(0).Value))
  ACTION(0).Caption = "-" + Trim(str(20 - VEL(0)))
 Case Is = 20
  Call PLAYERSETRATE("DECK1", 100)
  ACTION(0).Caption = "0"
 Case Is > 20
  Call PLAYERSETRATE("DECK1", 100 + (VEL(0).Value - 20))
  ACTION(0).Caption = "+" + Trim(str(VEL(0) - 20))
 End Select
End Sub

Private Sub VELOCC2()
 Select Case VEL(1)
 Case Is < 20
  Call PLAYERSETRATE("DECK2", 100 - (20 - VEL(1).Value))
  ACTION(1).Caption = "-" + Trim(str(20 - VEL(1)))
 Case Is = 20
  Call PLAYERSETRATE("DECK2", 100)
  ACTION(1).Caption = "0"
 Case Is > 20
  Call PLAYERSETRATE("DECK2", 100 + (VEL(1).Value - 20))
  ACTION(1).Caption = "+" + Trim(str(VEL(1) - 20))
 End Select
End Sub
