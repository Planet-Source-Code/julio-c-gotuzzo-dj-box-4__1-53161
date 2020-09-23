VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Dj Box 4 - Cargar Tema/s"
   ClientHeight    =   5760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7185
   ControlBox      =   0   'False
   Icon            =   "archiv.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox dirlist 
      Height          =   2400
      Left            =   0
      TabIndex        =   12
      Top             =   1650
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.ListBox listFiles 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   3765
      ItemData        =   "archiv.frx":0442
      Left            =   3765
      List            =   "archiv.frx":0444
      MultiSelect     =   2  'Extended
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   3210
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3570
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   -15
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   180
      TabIndex        =   2
      Top             =   5205
      Width           =   3405
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   4815
      Left            =   180
      TabIndex        =   1
      Top             =   195
      Width           =   3420
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   3795
      Hidden          =   -1  'True
      Left            =   3765
      MultiSelect     =   2  'Extended
      Pattern         =   "*.mp3"
      TabIndex        =   0
      Top             =   225
      Width           =   3225
   End
   Begin VB.Label btnGet 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar Archivos"
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   4620
      TabIndex        =   10
      Top             =   4755
      Width           =   1500
   End
   Begin VB.Label SELS 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ninguno"
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   2
      Left            =   5805
      TabIndex        =   9
      ToolTipText     =   "Reproducir."
      Top             =   4425
      Width           =   870
   End
   Begin VB.Label SELS 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Invertir"
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   1
      Left            =   4920
      TabIndex        =   8
      ToolTipText     =   "Reproducir."
      Top             =   4425
      Width           =   870
   End
   Begin VB.Label SELS 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Todos"
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   0
      Left            =   4020
      TabIndex        =   7
      ToolTipText     =   "Reproducir."
      Top             =   4425
      Width           =   870
   End
   Begin VB.Label labelResult 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0 Archivos."
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   3765
      TabIndex        =   6
      Top             =   4095
      Width           =   3210
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
      Height          =   270
      Left            =   5550
      TabIndex        =   4
      Top             =   5250
      Width           =   1215
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
      Height          =   270
      Left            =   3945
      TabIndex        =   3
      Top             =   5250
      Width           =   1215
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   5895
      Left            =   -30
      Picture         =   "archiv.frx":0446
      Top             =   -60
      Width           =   7290
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer

Private Sub btnGet_Click()
 Form2.Enabled = False
 If listFiles.Visible = False Then listFiles.Visible = True
 If File1.Visible = True Then File1.Visible = False
 Call BuscarArchivos(Dir1.path, "mp3", listFiles, dirlist, False)
 ordenar2 listFiles, dirlist, Indicador, False
 labelResult.Caption = listFiles.ListCount & " Archivos."
 Form2.Enabled = True
End Sub

Private Sub btnGet_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If btnGet.ForeColor <> SELCOLOR Then btnGet.ForeColor = SELCOLOR
 If Label1.ForeColor <> UNSELCOLOR Then Label1.ForeColor = UNSELCOLOR
 If SELS(0).ForeColor <> UNSELCOLOR Then SELS(0).ForeColor = UNSELCOLOR
 If SELS(1).ForeColor <> UNSELCOLOR Then SELS(1).ForeColor = UNSELCOLOR
 If SELS(2).ForeColor <> UNSELCOLOR Then SELS(2).ForeColor = UNSELCOLOR
 If Label2.ForeColor <> UNSELCOLOR Then Label2.ForeColor = UNSELCOLOR
End Sub

Private Sub Dir1_Change()
 File1.path = Dir1.path
 If listFiles.Visible = True Then listFiles.Visible = False
 If File1.Visible = False Then File1.Visible = True
 labelResult.Caption = File1.ListCount & " Archivos."
End Sub

Private Sub Dir1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Label2.ForeColor <> UNSELCOLOR Then Label2.ForeColor = UNSELCOLOR
 If Label1.ForeColor <> UNSELCOLOR Then Label1.ForeColor = UNSELCOLOR
End Sub

Private Sub Drive1_Change()
 Dir1.path = Drive1.Drive
End Sub

Private Sub File1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Label2.ForeColor <> UNSELCOLOR Then Label2.ForeColor = UNSELCOLOR
 If Label1.ForeColor <> UNSELCOLOR Then Label1.ForeColor = UNSELCOLOR
End Sub

Private Sub Form_Load()
 If SKINFILE > 0 Then Call SET_SKIN(4, SKINFILE)
 Dir1.ForeColor = DATAFORE
 Dir1.BackColor = DATABACK
 Drive1.ForeColor = DATAFORE
 Drive1.BackColor = DATABACK
 File1.ForeColor = DATAFORE
 File1.BackColor = DATABACK
 listFiles.ForeColor = DATAFORE
 listFiles.BackColor = DATABACK
 Label1.ForeColor = UNSELCOLOR
 Label2.ForeColor = UNSELCOLOR
 SELS(0).ForeColor = UNSELCOLOR
 SELS(1).ForeColor = UNSELCOLOR
 SELS(2).ForeColor = UNSELCOLOR
 btnGet.ForeColor = UNSELCOLOR
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Label2.ForeColor <> UNSELCOLOR Then Label2.ForeColor = UNSELCOLOR
 If Label1.ForeColor <> UNSELCOLOR Then Label1.ForeColor = UNSELCOLOR
 If SELS(0).ForeColor <> UNSELCOLOR Then SELS(0).ForeColor = UNSELCOLOR
 If SELS(1).ForeColor <> UNSELCOLOR Then SELS(1).ForeColor = UNSELCOLOR
 If SELS(2).ForeColor <> UNSELCOLOR Then SELS(2).ForeColor = UNSELCOLOR
 If btnGet.ForeColor <> UNSELCOLOR Then btnGet.ForeColor = UNSELCOLOR
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 n = 0
 If File1.Visible = True Then
  If File1.ListIndex > -1 And File1.ListCount > 0 And FILESELECTED(File1) > 0 Then
   Form2.Enabled = False
   If Text1 = 0 Then
    Call CargarTemas(Dir1.path, dirlist, File1, Form1.LISTA(0), Form1.LISTA(1), True, False, False, False, 0, False)
   Else
    Call CargarTemas(Dir1.path, dirlist, File1, Form1.LISTB(0), Form1.LISTB(1), True, False, False, False, 0, False)
   End If
   Indicador.Enabled = True
   Indicador.Visible = True
   Indicador.Caption = "Creando Dependencias..."
   If Text1 = 0 Then
    Indicador.ProgressBar1.Max = Form1.LISTA(0).ListCount
   Else
    Indicador.ProgressBar1.Max = Form1.LISTB(0).ListCount
   End If
   Indicador.ProgressBar1.Value = 0
   If Text = 0 Then
    n = 0
    Do While n <= Form1.LISTA(0).ListCount - 1
     DoEvents
     Form1.ILISTA.AddItem 0
     Form1.FLISTA.AddItem 0
     estate Indicador.Label1, "Generando"
     Indicador.ProgressBar1.Value = Indicador.ProgressBar1 + 1
     n = n + 1
    Loop
   Else
    n = 0
    Do While n <= Form1.LISTB(0).ListCount - 1
     DoEvents
     Form1.ILISTB.AddItem 0
     Form1.FLISTB.AddItem 0
     estate Indicador.Label1, "Generando"
     Indicador.ProgressBar1.Value = Indicador.ProgressBar1 + 1
     n = n + 1
    Loop
   End If
   Unload Indicador
   If Text1 = 0 Then
    If Form1.LISTA(0).ListCount > 0 And Form1.LISTA(0).ListIndex < 0 Then Form1.LISTA(0).ListIndex = 0
   Else
    If Form1.LISTB(0).ListCount > 0 And Form1.LISTB(0).ListIndex < 0 Then Form1.LISTB(0).ListIndex = 0
   End If
   Form2.Enabled = True
   Form1.Enabled = True
   Form4.Enabled = True
   Form5.Enabled = True
   Unload Me
  Else
   MsgBox ("Debe Seleccionar Un Archivo De Audio")
  End If
 Else
  If listFiles.ListIndex > -1 And listFiles.ListCount > 0 And listFiles.SelCount > 0 Then
   Form2.Enabled = False
   If Text1 = 0 Then
    Call CargarTemas(Dir1.path, dirlist, listFiles, Form1.LISTA(0), Form1.LISTA(1), True, True, False, False, 0, True)
   Else
    Call CargarTemas(Dir1.path, dirlist, listFiles, Form1.LISTB(0), Form1.LISTB(1), True, True, False, False, 0, True)
   End If
   Indicador.Enabled = True
   Indicador.Visible = True
   Indicador.Caption = "Creando Dependencias..."
   If Text1 = 0 Then
    Indicador.ProgressBar1.Max = Form1.LISTA(0).ListCount
   Else
    Indicador.ProgressBar1.Max = Form1.LISTB(0).ListCount
   End If
   Indicador.ProgressBar1.Value = 0
   If Text = 0 Then
    n = 0
    Do While n <= Form1.LISTA(0).ListCount - 1
     DoEvents
     Form1.ILISTA.AddItem 0
     Form1.FLISTA.AddItem 0
     estate Indicador.Label1, "Generando"
     Indicador.ProgressBar1.Value = Indicador.ProgressBar1 + 1
     n = n + 1
    Loop
   Else
    n = 0
    Do While n <= Form1.LISTB(0).ListCount - 1
     DoEvents
     Form1.ILISTB.AddItem 0
     Form1.FLISTB.AddItem 0
     estate Indicador.Label1, "Generando"
     Indicador.ProgressBar1.Value = Indicador.ProgressBar1 + 1
     n = n + 1
    Loop
   End If
   Unload Indicador
   If Text1 = 0 Then
    If Form1.LISTA(0).ListCount > 0 And Form1.LISTA(0).ListIndex < 0 Then Form1.LISTA(0).ListIndex = 0
   Else
    If Form1.LISTB(0).ListCount > 0 And Form1.LISTB(0).ListIndex < 0 Then Form1.LISTB(0).ListIndex = 0
   End If
   Form2.Enabled = True
   Form1.Enabled = True
   Form4.Enabled = True
   Form5.Enabled = True
   Unload Me
  Else
   MsgBox ("Debe Seleccionar Un Archivo De Audio")
  End If
 End If
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Label1.ForeColor <> SELCOLOR Then Label1.ForeColor = SELCOLOR
 If Label2.ForeColor <> UNSELCOLOR Then Label2.ForeColor = UNSELCOLOR
 If SELS(0).ForeColor <> UNSELCOLOR Then SELS(0).ForeColor = UNSELCOLOR
 If SELS(1).ForeColor <> UNSELCOLOR Then SELS(1).ForeColor = UNSELCOLOR
 If SELS(2).ForeColor <> UNSELCOLOR Then SELS(2).ForeColor = UNSELCOLOR
 If btnGet.ForeColor <> UNSELCOLOR Then btnGet.ForeColor = UNSELCOLOR
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Form1.Enabled = True
 Form4.Enabled = True
 Form5.Enabled = True
 Unload Me
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Label2.ForeColor <> SELCOLOR Then Label2.ForeColor = SELCOLOR
 If Label1.ForeColor <> UNSELCOLOR Then Label1.ForeColor = UNSELCOLOR
 If SELS(0).ForeColor <> UNSELCOLOR Then SELS(0).ForeColor = UNSELCOLOR
 If SELS(1).ForeColor <> UNSELCOLOR Then SELS(1).ForeColor = UNSELCOLOR
 If SELS(2).ForeColor <> UNSELCOLOR Then SELS(2).ForeColor = UNSELCOLOR
 If btnGet.ForeColor <> UNSELCOLOR Then btnGet.ForeColor = UNSELCOLOR
End Sub

Private Sub SELS_Click(Index As Integer)
 Form2.MousePointer = 11
 Select Case Index
 Case Is = 0
  n = 0
  If File1.Visible = True Then
   Do While n <= File1.ListCount - 1
    If File1.SELECTED(n) = False Then File1.SELECTED(n) = True
    n = n + 1
   Loop
  Else
   Do While n <= listFiles.ListCount - 1
    If listFiles.SELECTED(n) = False Then listFiles.SELECTED(n) = True
    n = n + 1
   Loop
  End If
 Case Is = 1
  n = 0
  If File1.Visible = True Then
   Do While n <= File1.ListCount - 1
    If File1.SELECTED(n) = False Then File1.SELECTED(n) = True Else File1.SELECTED(n) = False
    n = n + 1
   Loop
  Else
   If listFiles.SelCount > 0 Then
    Do While n <= listFiles.ListCount - 1
     If listFiles.SELECTED(n) = False Then listFiles.SELECTED(n) = True Else listFiles.SELECTED(n) = False
     n = n + 1
    Loop
   End If
  End If
 Case Is = 2
  n = 0
  If File1.Visible = True Then
   Do While n <= File1.ListCount - 1
    If File1.SELECTED(n) = True Then File1.SELECTED(n) = False
    n = n + 1
   Loop
  Else
   If listFiles.SelCount > 0 Then
    Do While n <= listFiles.ListCount - 1
     If listFiles.SELECTED(n) = True Then listFiles.SELECTED(n) = False
     n = n + 1
    Loop
   End If
  End If
 End Select
 Form2.MousePointer = 0
End Sub

Private Sub Timer1_Timer()
 If File1.Visible = True And labelResult.Caption <> File1.ListCount & " Archivos." Then labelResult.Caption = File1.ListCount & " Archivos."
End Sub

Private Sub SELS_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 Select Case Index
 Case Is = 0
  If SELS(0).ForeColor <> SELCOLOR Then SELS(0).ForeColor = SELCOLOR
  If SELS(1).ForeColor <> UNSELCOLOR Then SELS(1).ForeColor = UNSELCOLOR
  If SELS(2).ForeColor <> UNSELCOLOR Then SELS(2).ForeColor = UNSELCOLOR
 Case Is = 1
  If SELS(1).ForeColor <> SELCOLOR Then SELS(1).ForeColor = SELCOLOR
  If SELS(0).ForeColor <> UNSELCOLOR Then SELS(0).ForeColor = UNSELCOLOR
  If SELS(2).ForeColor <> UNSELCOLOR Then SELS(2).ForeColor = UNSELCOLOR
 Case Is = 2
  If SELS(2).ForeColor <> SELCOLOR Then SELS(2).ForeColor = SELCOLOR
  If SELS(0).ForeColor <> UNSELCOLOR Then SELS(0).ForeColor = UNSELCOLOR
  If SELS(1).ForeColor <> UNSELCOLOR Then SELS(1).ForeColor = UNSELCOLOR
 End Select
 If Label1.ForeColor <> UNSELCOLOR Then Label1.ForeColor = UNSELCOLOR
 If Label2.ForeColor <> UNSELCOLOR Then SELS(0).ForeColor = UNSELCOLOR
 If btnGet.ForeColor <> UNSELCOLOR Then btnGet.ForeColor = UNSELCOLOR
End Sub

Private Function GETTOTALSIZE(List As FileListBox) As Integer
 Dim pp As Integer
  pp = 0
  GETTOTALSIZE = 0
  Do While pp <= List.ListCount - 1
   If List.SELECTED(pp) = True Then
    Call getMP3Info(Dir1.path + "\" + List.List(pp), MP3INF)
    GETTOTALSIZE = GETTOTALSIZE + MP3INF.LENGTH
   End If
   pp = pp + 1
  Loop
End Function
