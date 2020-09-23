Attribute VB_Name = "PLS_Loader"
Public Function PLSSAVE(Filename As String, Namelist As ListBox, Filelist As ListBox, SELECTED As Boolean) As Boolean
 Dim i As Integer
 PLSSAVE = False
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
  Print #2, "[playlist]"
  For i = 0 To Namelist.ListCount - 1
    If SELECTED = True Then
     If Namelist.SELECTED(i) = True Then
       Print #2, "File" + Trim(str(i + 1)) + "=" + Trim(dirlist.List(i)) + Trim(Filelist.List(i))
       Print #2, "Title" + Trim(str(i + 1)) + "=" + Trim(Namelist.List(i))
       If getMP3Info(dirlist.List(i) + Filelist.List(i), MP3INF) = True Then
        Print #2, "Length" + Trim(str(i + 1)) + "=" + Mid(Trim(MP3INF.LENGTH), 1, Len(Trim(MP3INF.LENGTH)) - 8)
       Else
        Print #2, "Length" + Trim(str(i + 1)) + "=0"
       End If
      Indicador.ProgressBar1.Value = Indicador.ProgressBar1.Value + 1
     End If
    Else
      Print #2, "File" + Trim(str(i + 1)) + "=" + Trim(dirlist.List(i)) + Trim(Filelist.List(i))
      Print #2, "Title" + Trim(str(i + 1)) + "=" + Trim(Namelist.List(i))
      If getMP3Info(dirlist.List(i) + Filelist.List(i), MP3INF) = True Then
       Print #2, "Length" + Trim(str(i + 1)) + "=" + Mid(Trim(MP3INF.LENGTH), 1, Len(Trim(MP3INF.LENGTH)) - 8)
      Else
       Print #2, "Length" + Trim(str(i + 1)) + "=0"
      End If
     Indicador.ProgressBar1.Value = Indicador.ProgressBar1.Value + 1
    End If
  Next i
  Print #2, "NumberOfEntries =" + Trim(str(Namelist.ListCount))
  Print #2, "Version = 2"
  PLSSAVE = True
 End If
err:
 Close #2
 Unload Indicador
End Function

Public Function PLSLOAD(Filename As String, Namelist As ListBox, Filelist As ListBox) As Boolean
 Dim tempor As String
 tempor = ""
 Open Filename For Input As #1
  Line Input #1, a$
  If a$ <> "[playlist]" Then
   MsgBox ("ERROR: Este Archivo No Es Una Lista PLS")
   Exit Function
  End If
 Namelist.Clear
 Filelist.Clear
 Do
  Line Input #1, a$
  If Trim(UponSimbol(a$, "=", 0)) = "NumberOfEntries" Then Exit Do
  If InStr(1, UponSimbol(a$, "=", 1), "\", vbBinaryCompare) <= 0 Then
   tempor = left(Filename, InStrRev(Filename, "\", , vbTextCompare)) + UponSimbol(a$, "=", 1)
  Else
   tempor = UponSimbol(a$, "=", 1)
  End If
  If Dir(tempor, vbArchive) <> "" Then
   Filelist.AddItem tempor
  Else
   Prev.List1.AddItem tempor
  End If
  Line Input #1, a$
  If Dir(tempor, vbArchive) <> "" Then Namelist.AddItem UponSimbol(a$, "=", 1)
  Line Input #1, a$
 Loop While Not (EOF(1))
 Close #1
 If Prev.List1.ListCount > 0 Then
  Prev.Caption = "Estos archivos no existen"
  Prev.Enabled = True
  Prev.Visible = True
 End If
End Function

