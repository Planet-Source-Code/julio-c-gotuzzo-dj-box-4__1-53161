Attribute VB_Name = "M3U_Loader"
Public Function M3ULOAD(Filename As String, Namelist As ListBox, Filelist As ListBox)
 Dim tempor As String
 tempor = ""
 Open Filename For Input As #1
  Line Input #1, a$
  If a$ <> "#EXTM3U" Then
   MsgBox ("ERROR: Este Archivo No Es Una Lista M3U")
   Exit Function
  End If
 Namelist.Clear
 Filelist.Clear
 Do
  Line Input #1, a$
  tempor = UponSimbol(a$, ",", 1)
  Line Input #1, a$
  If InStr(1, a$, "\", vbBinaryCompare) <= 0 Then
   If Dir(left(Filename, InStrRev(Filename, "\", , vbTextCompare)) + a$, vbArchive) <> "" Then
    Filelist.AddItem left(Filename, InStrRev(Filename, "\", , vbTextCompare)) + a$
    Namelist.AddItem tempor
   Else
    Prev.List1.AddItem left(Filename, InStrRev(Filename, "\", , vbTextCompare)) + a$
   End If
  Else
   If Dir(a$, vbArchive) <> "" Then
    Filelist.AddItem a$
    Namelist.AddItem tempor
   Else
    Prev.List1.AddItem a$
   End If
  End If
 Loop While Not (EOF(1))
 Close #1
 If Prev.List1.ListCount > 0 Then
  Prev.Caption = "Estos archivos no existen"
  Prev.Enabled = True
  Prev.Visible = True
 End If
End Function

Public Function M3USAVE(Filename As String, Namelist As ListBox, Filelist As ListBox, dirlist As ListBox, SELECTED As Boolean, Usardirlist As Boolean) As Boolean
 Dim i As Integer
 M3USAVE = False
 Indicador.Visible = True
 Indicador.Enabled = True
 If SELECTED = True Then
  Indicador.ProgressBar1.Max = Namelist.SelCount
 Else
  Indicador.ProgressBar1.Max = Namelist.ListCount
 End If
 Indicador.ProgressBar1.Value = 0
 Indicador.Label1.Caption = "Guardando: " + Trim(Filename) + "..."
 On Error GoTo err
 If Namelist.ListCount > 0 Then
  Open Filename For Output As #2
  Print #2, "#EXTM3U"
  For i = 0 To Namelist.ListCount - 1
   If SELECTED = True Then
    If Namelist.SELECTED(i) = True Then
     If Usardirlist = False Then
      If getMP3Info(SOLOPATH(Filename) + Filelist.List(i), MP3INF) = True Then
       Print #2, "#EXTINF:" + Mid(Trim(MP3INF.LENGTH), 1, Len(Trim(MP3INF.LENGTH)) - 8) + "," + Trim(Namelist.List(i))
      Else
       Print #2, "#EXTINF:0," + Trim(Namelist.List(i))
      End If
      Print #2, Trim(SOLOPATH(Filename)) + Trim(Filelist.List(i))
     Else
      If getMP3Info(dirlist.List(i) + Filelist.List(i), MP3INF) = True Then
       Print #2, "#EXTINF:" + Mid(Trim(MP3INF.LENGTH), 1, Len(Trim(MP3INF.LENGTH)) - 8) + "," + Trim(Namelist.List(i))
      Else
       Print #2, "#EXTINF:0," + Trim(Namelist.List(i))
      End If
      Print #2, Trim(dirlist.List(i)) + Trim(Filelist.List(i))
     End If
     Indicador.ProgressBar1.Value = Indicador.ProgressBar1.Value + 1
    End If
   Else
    If Usardirlist = False Then
     If getMP3Info(SOLOPATH(Filename) + Filelist.List(i), MP3INF) = True Then
      Print #2, "#EXTINF:" + Mid(Trim(MP3INF.LENGTH), 1, Len(Trim(MP3INF.LENGTH)) - 8) + "," + Trim(Namelist.List(i))
     Else
      Print #2, "#EXTINF:0," + Trim(Namelist.List(i))
     End If
     Print #2, Trim(SOLOPATH(Filename)) + Trim(Filelist.List(i))
    Else
     If getMP3Info(dirlist.List(i) + Filelist.List(i), MP3INF) = True Then
      Print #2, "#EXTINF:" + Mid(Trim(MP3INF.LENGTH), 1, Len(Trim(MP3INF.LENGTH)) - 8) + "," + Trim(Namelist.List(i))
     Else
      Print #2, "#EXTINF:0," + Trim(Namelist.List(i))
     End If
     Print #2, Trim(dirlist.List(i)) + Trim(Filelist.List(i))
    End If
    Indicador.ProgressBar1.Value = Indicador.ProgressBar1.Value + 1
   End If
  Next i
  M3USAVE = True
 End If
err:
 Close #2
 MsgBox ("Se Produjo Un Error Al Intentar Guardar La Lista")
 Unload Indicador
End Function
