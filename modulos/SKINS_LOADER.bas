Attribute VB_Name = "SKINS_LOADER"
Function SKINS_LIST(Combo As ComboBox) As Boolean
 Dim check As Boolean
 Dim POSIT As Integer
 Dim N As Integer
 check = False
 POSIT = 0
 SKINS_LIST = False
 Open App.path & "\Skins\skins.dat" For Input As #1
 Do
  Line Input #1, a$
  If check = False And a$ = "#SKINS_HEADER" Then
   Combo.Clear
   Combo.AddItem "Default."
   check = True
   POSIT = 0
   N = 1
   SKINS_LIST = True
  End If
  If check = False Then
   Exit Do
  Else
   If a$ = "[FIN]" Then Exit Do
   If a$ = "[" & N & "]" Then
    POSIT = 1
   Else
    If POSIT = 1 Then
     Combo.AddItem right(a$, Len(a$) - 7)
     POSIT = 0
     N = N + 1
    End If
   End If
  End If
 Loop While Not (EOF(1))
 Close #1
End Function

Function SET_SKIN(CONSOLE As Integer, SKINNUMBER As Integer) As Boolean
 Dim check As Boolean
 Dim POSIT As Integer
 check = False
 POSIT = 0
 SET_SKIN = False
 Open App.path & "\Skins\skins.dat" For Input As #1
 Do
  Line Input #1, a$
  If check = False And a$ = "#SKINS_HEADER" Then
   check = True
   POSIT = 0
   SET_SKIN = True
  End If
  If check = False Then
   Exit Do
  Else
   If a$ = "[FIN]" Then Exit Do
   If a$ = "[" & SKINNUMBER & "]" Then
    POSIT = 1
   Else
    If InStr(a$, "Ruta=") > 0 Then
     Select Case CONSOLE
     Case Is = 0
       Form1.Image1.Picture = LoadPicture(App.path & "\SKINS\" + right(a$, Len(a$) - 5) + "player.bmp")
       Exit Do
     Case Is = 1
       Form5.Image2.Picture = LoadPicture(App.path & "\SKINS\" + right(a$, Len(a$) - 5) + "mixer.bmp")
       Exit Do
     Case Is = 2
       Form4.image3.Picture = LoadPicture(App.path & "\SKINS\" + right(a$, Len(a$) - 5) + "fx.bmp")
       Exit Do
     Case Is = 3
       Form8.Image1.Picture = LoadPicture(App.path & "\SKINS\" + right(a$, Len(a$) - 5) + "cdplayer.bmp")
       Exit Do
     Case Is = 4
       Form2.Image1.Picture = LoadPicture(App.path & "\SKINS\" + right(a$, Len(a$) - 5) + "file1.bmp")
       Exit Do
     Case Is = 5
       Form3.Image1.Picture = LoadPicture(App.path & "\SKINS\" + right(a$, Len(a$) - 5) + "file2.bmp")
       Exit Do
     Case Is = 6
       Form6.Image1.Picture = LoadPicture(App.path & "\SKINS\" + right(a$, Len(a$) - 5) + "buscar.bmp")
       Exit Do
     End Select
    End If
   End If
  End If
 Loop While Not (EOF(1))
 Close #1
 SET_SKIN = True
End Function
