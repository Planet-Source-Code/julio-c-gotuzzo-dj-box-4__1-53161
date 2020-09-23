Attribute VB_Name = "Player_Mod"
Public Sub CargarTemas(Directorio As String, Dires As ListBox, ArchOrigen As Object, NombreDest As ListBox, ArchDest As ListBox, Ordenar As Boolean, Comprobar As Boolean, Numerar As Boolean, Filtrar As Boolean, TipoFilt As Integer, UsarDires As Boolean)
 Dim n As Integer
 n = 0
 NombreDest.Clear
 ArchDest.Clear
 Indicador.Enabled = True
 Indicador.Visible = True
 Indicador.ProgressBar1.Max = ArchOrigen.ListCount
 Indicador.ProgressBar1.Value = 0
 Do While n <= ArchOrigen.ListCount - 1
  DoEvents
  If UsarDires = True Then Directorio = Dires.List(n)
  If Mid(Directorio, Len(Directorio), 1) <> "\" And Mid(Directorio, Len(Directorio), 1) <> ":" Then Directorio = Directorio + "\"
  id3info.Album = ""
  id3info.Artist = ""
  If FileLen(Directorio + ArchOrigen.List(n)) > 0 And GetId3(Directorio + ArchOrigen.List(n)) = True And Trim(id3info.Artist) <> "" And Trim(id3info.Title) <> "" Then
   If Filtrar = False Or Filtrar = True And TipoFilt <> 1 Then
    If Trim(id3info.Artist) <> "" Then
     If Numerar = True Then
      If Ordenar = False Then
       NombreDest.AddItem Trim(str(NombreDest.ListCount + 1)) + ". " + SPELL(Trim(id3info.Artist))
      Else
       NombreDest.AddItem SPELL(Trim(id3info.Artist))
      End If
     Else
      NombreDest.AddItem SPELL(Trim(id3info.Artist))
     End If
     If Trim(id3info.Title) <> "" Then NombreDest.List(NombreDest.ListCount - 1) = NombreDest.List(NombreDest.ListCount - 1) + " - " + SPELL(Trim(id3info.Title))
    Else
     If Numerar = True Then
      If Trim(id3info.Title) <> "" Then
       If Ordenar = False Then
        NombreDest.AddItem Trim(str(NombreDest.ListCount + 1)) + ". " + SPELL(Trim(id3info.Title))
       Else
        NombreDest.AddItem SPELL(Trim(id3info.Title))
       End If
      End If
     Else
      If Trim(id3info.Title) <> "" Then NombreDest.AddItem SPELL(Trim(id3info.Title))
     End If
    End If
    If UsarDires = False Then
     ArchDest.AddItem Directorio + ArchOrigen.List(n)
    Else
     ArchDest.AddItem Dires.List(n) + ArchOrigen.List(n)
    End If
   End If
  Else
   If Filtrar = False Or Filtrar = True And TipoFilt <> 0 Then
    If Numerar = True Then
     If Ordenar = False Then
      NombreDest.AddItem Trim(str(NombreDest.ListCount + 1)) + ". " + SPELL(Trim(SINEXT(SOLOFILE(ArchOrigen.List(n))))) + " (Sin Id3Tag)"
     Else
      NombreDest.AddItem SPELL(Trim(SINEXT(SOLOFILE(ArchOrigen.List(n))))) + " (Sin Id3Tag)"
     End If
    Else
     NombreDest.AddItem SPELL(Trim(SINEXT(SOLOFILE(ArchOrigen.List(n))))) + " (Sin Id3Tag)"
    End If
    If UsarDires = False Then
     ArchDest.AddItem Directorio + ArchOrigen.List(n)
    Else
     ArchDest.AddItem Dires.List(n) + ArchOrigen.List(n)
    End If
   End If
  End If
  Indicador.ProgressBar1.Value = Indicador.ProgressBar1.Value + 1
  Indicador.Label1 = "Buscando Id3Tag: " + Trim(ArchOrigen.List(n)) + "..."
  n = n + 1
 Loop
 Unload Indicador
 n = 0
 Do While n <= NombreDest.ListCount - 1
  If Trim(NombreDest.List(n)) = "-" Then
   NombreDest.RemoveItem n
   ArchDest.RemoveItem n
   n = 0
  End If
  n = n + 1
 Loop
 If Numerar = True Then
  If Ordenar = True Then ordenar2 NombreDest, ArchDest, Indicador, True
 Else
  If Ordenar = True Then ordenar2 NombreDest, ArchDest, Indicador, False
 End If
End Sub

Public Sub ordenar2(lista1 As ListBox, lista2 As ListBox, Indicador As Form, Numerar As Boolean)
 Dim n As Integer
 Dim f As Integer
 Dim min As String
 Dim pos As Integer
 Dim total As Integer
 Dim total2 As Integer
 If lista1.ListCount < 1 Then Exit Sub
 If lista1.ListCount <> lista2.ListCount Then Exit Sub
 Indicador.Enabled = True
 Indicador.Visible = True
 Indicador.ProgressBar1.Max = lista1.ListCount
 Indicador.ProgressBar1.Value = 0
 total = lista1.ListCount
 total2 = lista1.ListCount + 1
 n = 1
 min = "zzzzzzz"
 pos = 0
 Do While n <= total
    DoEvents
  f = 0
  min = "zzzzzzz"
  total2 = total2 - 1
  Do While f <= total2 - 1
   If lista1.List(f) < min Then
    pos = f
    min = lista1.List(f)
   End If
   f = f + 1
  Loop
  If Numerar = True Then
   lista1.AddItem Trim(str(n)) + ". " + lista1.List(pos)
  Else
   lista1.AddItem lista1.List(pos)
  End If
  lista2.AddItem lista2.List(pos)
  lista1.RemoveItem pos
  lista2.RemoveItem pos
  n = n + 1
  Indicador.Label1 = "Ordenando: " + Trim(lista1.List(pos)) + "..."
  Indicador.ProgressBar1.Value = Indicador.ProgressBar1.Value + 1
 Loop
 Unload Indicador
End Sub

Public Function temasdur(Lista As Object) As Long
 Dim n As Integer
 n = 0
 temasdur = 0
 Do While n <= Lista.ListCount - 1
  Call getMP3Info(Lista.List(n), MP3INF)
  temasdur = temasdur + Val(MP3INF.LENGTH)
  n = n + 1
 Loop
End Function
