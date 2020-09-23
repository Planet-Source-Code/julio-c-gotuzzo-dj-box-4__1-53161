Attribute VB_Name = "StringModule"
Option Explicit

Private Declare Function SendMessageA Lib _
"user32" (ByVal hWnd As Long, ByVal wMsg As Integer, ByVal _
wParam As Integer, lParam As Any) As Long

Private Const LB_FINDSTRING = &H18F

Public Function intbool(Valor As Integer) As Boolean
 If Valor = 0 Then
  intbool = False
 Else
  If Valor >= 1 Then intbool = True
 End If
End Function

Public Function DepurarUbi(Ubicacion As String) As String
 Ubicacion = Replace(Ubicacion, "\\", "\", 1, -1, vbTextCompare)
 If Len(Ubicacion) = 3 And Mid(Ubicacion, 2, 2) = ":\" Then Ubicacion = Mid(Ubicacion, 1, 2)
 If Len(Ubicacion) = 1 And Asc(Trim(UCase(Ubicacion))) >= 65 And Asc(Trim(UCase(Ubicacion))) <= 90 Then Ubicacion = Ubicacion + ":"
 DepurarUbi = Ubicacion
End Function

Public Function AgregarSigno(Directorio As String) As String
 If Mid(Directorio, Len(Directorio), 1) <> "\" Then AgregarSigno = Directorio + "\"
End Function

Public Function MINSEG(SEGUNDOS As Long) As String
 If Len(Trim(str(Int(SEGUNDOS / 60)))) = 1 And Len(Trim(str(Int(SEGUNDOS) - (Int(SEGUNDOS / 60) * 60)))) = 1 Then MINSEG = "0" + Trim(str(Int(SEGUNDOS / 60))) + ":0" + Trim(str(Int(SEGUNDOS) - (Int(SEGUNDOS / 60) * 60)))
 If Len(Trim(str(Int(SEGUNDOS / 60)))) = 2 And Len(Trim(str(Int(SEGUNDOS) - (Int(SEGUNDOS / 60) * 60)))) = 2 Then MINSEG = Trim(str(Int(SEGUNDOS / 60))) + ":" + Trim(str(Int(SEGUNDOS) - (Int(SEGUNDOS / 60) * 60)))
 If Len(Trim(str(Int(SEGUNDOS / 60)))) = 1 And Len(Trim(str(Int(SEGUNDOS) - (Int(SEGUNDOS / 60) * 60)))) = 2 Then MINSEG = "0" + Trim(str(Int(SEGUNDOS / 60))) + ":" + Trim(str(Int(SEGUNDOS) - (Int(SEGUNDOS / 60) * 60)))
 If Len(Trim(str(Int(SEGUNDOS / 60)))) = 2 And Len(Trim(str(Int(SEGUNDOS) - (Int(SEGUNDOS / 60) * 60)))) = 1 Then MINSEG = Trim(str(Int(SEGUNDOS / 60))) + ":0" + Trim(str(Int(SEGUNDOS) - (Int(SEGUNDOS / 60) * 60)))
End Function

Public Function HORMINSEG(SEGUNDOS As Long) As String
 If Len(Trim(str(Int(SEGUNDOS / 3600)))) = 1 And Len(Trim(str(Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) / 60)))) = 1 And Len(Trim(str(Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) - (Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) / 60) * 60))))) = 1 Then HORMINSEG = "0" + Trim(str(Int(SEGUNDOS / 3600))) + ":0" + Trim(str(Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) / 60))) + ":0" + Trim(str(Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) - (Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) / 60) * 60))))
 If Len(Trim(str(Int(SEGUNDOS / 3600)))) = 2 And Len(Trim(str(Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) / 60)))) = 2 And Len(Trim(str(Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) - (Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) / 60) * 60))))) = 2 Then HORMINSEG = Trim(str(Int(SEGUNDOS / 3600))) + ":" + Trim(str(Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) / 60))) + ":" + Trim(str(Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) - (Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) / 60) * 60))))
 If Len(Trim(str(Int(SEGUNDOS / 3600)))) = 1 And Len(Trim(str(Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) / 60)))) = 2 And Len(Trim(str(Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) - (Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) / 60) * 60))))) = 2 Then HORMINSEG = "0" + Trim(str(Int(SEGUNDOS / 3600))) + ":" + Trim(str(Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) / 60))) + ":" + Trim(str(Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) - (Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) / 60) * 60))))
 If Len(Trim(str(Int(SEGUNDOS / 3600)))) = 2 And Len(Trim(str(Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) / 60)))) = 1 And Len(Trim(str(Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) - (Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) / 60) * 60))))) = 1 Then HORMINSEG = Trim(str(Int(SEGUNDOS / 3600))) + ":0" + Trim(str(Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) / 60))) + ":0" + Trim(str(Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) - (Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) / 60) * 60))))
 If Len(Trim(str(Int(SEGUNDOS / 3600)))) = 1 And Len(Trim(str(Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) / 60)))) = 1 And Len(Trim(str(Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) - (Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) / 60) * 60))))) = 2 Then HORMINSEG = "0" + Trim(str(Int(SEGUNDOS / 3600))) + ":0" + Trim(str(Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) / 60))) + ":" + Trim(str(Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) - (Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) / 60) * 60))))
 If Len(Trim(str(Int(SEGUNDOS / 3600)))) = 2 And Len(Trim(str(Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) / 60)))) = 2 And Len(Trim(str(Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) - (Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) / 60) * 60))))) = 1 Then HORMINSEG = Trim(str(Int(SEGUNDOS / 3600))) + ":" + Trim(str(Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) / 60))) + ":0" + Trim(str(Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) - (Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) / 60) * 60))))
 If Len(Trim(str(Int(SEGUNDOS / 3600)))) = 1 And Len(Trim(str(Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) / 60)))) = 2 And Len(Trim(str(Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) - (Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) / 60) * 60))))) = 1 Then HORMINSEG = "0" + Trim(str(Int(SEGUNDOS / 3600))) + ":" + Trim(str(Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) / 60))) + ":0" + Trim(str(Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) - (Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) / 60) * 60))))
 If Len(Trim(str(Int(SEGUNDOS / 3600)))) = 2 And Len(Trim(str(Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) / 60)))) = 1 And Len(Trim(str(Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) - (Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) / 60) * 60))))) = 2 Then HORMINSEG = Trim(str(Int(SEGUNDOS / 3600))) + ":0" + Trim(str(Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) / 60))) + ":" + Trim(str(Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) - (Int((Int(SEGUNDOS) - (Int(SEGUNDOS / 3600) * 3600)) / 60) * 60))))
End Function

Public Function PISTA(Numero As Integer, LARGO As Integer) As String
 PISTA = String(LARGO - Len(Trim(str(Numero))), "0") + Trim(str(Numero))
End Function

Public Function SINEXT(Filename As String) As String
 SINEXT = Left(Filename, Len(Filename) - 4)
End Function

Public Function SOLOEXT(Filename As String) As String
 SOLOEXT = Right(Filename, 3)
End Function

Public Function SOLOFILE(Filename As String) As String
 SOLOFILE = Right(Filename, Len(Filename) - InStrRev(Filename, "\", , vbTextCompare))
End Function

Public Function SOLOPATH(Filename As String) As String
 SOLOPATH = Left(Filename, InStrRev(Filename, "\", , vbTextCompare))
End Function

Public Function SOLODRIVE(Filename As String)
 SOLODRIVE = Left(Filename, 2)
End Function

Public Function LISTUNIQUE(lst As Object, str As String) As Boolean
 Dim nIndex As Integer
 Dim bFound As Boolean
 With lst
 For nIndex = 0 To .ListCount - 1
  If .List(nIndex) = str Then
   bFound = True
   Exit For
  End If
  Next
  If Not bFound Then
   LISTUNIQUE = True
  Else
   LISTUNIQUE = False
  End If
 End With
End Function

Public Function calcsort(LISTA As ListBox, indice As String) As Integer
 Dim f As Integer
 Dim min As String
 If LISTA.ListCount < 1 Or indice = "" Then Exit Function
 min = "zzzzzzz"
 f = 0
 calcsort = 0
 Do While f <= LISTA.ListCount - 1
  If LISTA.List(f) < min And LISTA.List(f) > indice Then
   calcsort = f
   min = LISTA.List(f)
  End If
  f = f + 1
 Loop
End Function

Public Function SPELL(Texto As String) As String
 Dim n As Integer
 n = 1
 Do While n <= Len(Texto)
  If n = 1 Then
   SPELL = UCase(Mid(Texto, 1, 1))
  Else
   If Mid(Texto, n - 1, 1) = " " Then
    SPELL = SPELL + UCase(Mid(Texto, n, 1))
   Else
    SPELL = SPELL + LCase(Mid(Texto, n, 1))
   End If
  End If
  n = n + 1
 Loop
 SPELL = Trim(SPELL)
End Function

Public Function BUSCARENLISTA(List As ListBox, itemtext As String) As String
  BUSCARENLISTA = SendMessageA(List.hWnd, _
  LB_FINDSTRING, -1, ByVal itemtext)
End Function

Public Function FILESELECTED(FileBox As FileListBox) As Integer
 Dim n As Integer
 n = 0
 FILESELECTED = 0
 Do While n <= FileBox.ListCount - 1
  If FileBox.SELECTED(n) = True Then FILESELECTED = FILESELECTED + 1
  n = n + 1
 Loop
End Function

Public Function ANTES(Texto As String, SEPARADOR As String) As String
 If Len(SEPARADOR) > 1 Or Trim(Texto) = "" Then Exit Function
 ANTES = Trim(Mid(Texto, 1, InStr(1, Trim(Texto), Trim(SEPARADOR), vbTextCompare) - 1))
End Function

Public Function DESPUES(Texto As String, SEPARADOR As String) As String
 If Len(SEPARADOR) > 1 Or Trim(Texto) = "" Then Exit Function
 DESPUES = Trim(Mid(Texto, InStr(1, Trim(Texto), Trim(SEPARADOR), vbTextCompare) + 1, Len(Trim(Texto)) - (InStr(1, Trim(Texto), Trim(SEPARADOR), vbTextCompare))))
End Function

Public Sub estate(Contenedor As Object, Texto As String)
 If InStr(1, LCase(Trim(Contenedor)), Texto, vbTextCompare) = 0 Then Contenedor = Texto + " /"
 If InStr(1, Contenedor, " |", vbTextCompare) <> 0 Then
  Contenedor.Caption = Replace(Trim(Contenedor.Caption), " |", " /", 1, -1, vbTextCompare)
 Else
  If InStr(1, Contenedor, " /", vbTextCompare) <> 0 Then
   Contenedor.Caption = Replace(Trim(Contenedor.Caption), " /", " -", 1, -1, vbTextCompare)
  Else
   If InStr(1, Contenedor, " -", vbTextCompare) <> 0 Then
    Contenedor.Caption = Replace(Trim(Contenedor.Caption), " -", " \", 1, -1, vbTextCompare)
   Else
    If InStr(1, Contenedor, " \", vbTextCompare) <> 0 Then
     Contenedor.Caption = Replace(Trim(Contenedor.Caption), " \", " |", 1, -1, vbTextCompare)
    End If
   End If
  End If
 End If
End Sub

Function Punto(Valor As String) As String
 Dim temp As String
 Dim conter As Integer
 Dim n As Integer
 n = Len(Valor)
 conter = 0
 Do While n >= 1
  DoEvents
  temp = temp + Mid(Valor, n, 1)
  conter = conter + 1
  If conter = 3 Then
   conter = 0
   temp = temp + "."
  End If
  n = n - 1
 Loop
 Punto = StrReverse(temp)
End Function

Public Function SOLOMEGAS(Valor As Long) As String
 Select Case Len(Trim(str(Valor)))
 Case Is <= 6
  SOLOMEGAS = Mid(str(Valor / 1048576), 1, 4)
 Case Is > 6
 End Select
End Function

Public Function SELITEMS(LISTA As Object) As Long
 Dim n As Integer
 SELITEMS = 0
 n = 0
 Do While n <= LISTA.ListCount - 1
  If LISTA.SELECTED(n) = True Then SELITEMS = SELITEMS + 1
  n = n + 1
 Loop
End Function

Public Sub BuscarArchivos(RutaInicial As String, Tipo As String, DestinoArch As ListBox, DestinoDir As ListBox, unir As Boolean)
 Dim txtfile As String
 Dim Y As Integer
 Dim X As Integer
 Dim path As String
 Dim tfilename As String
 Busqueda.Enabled = True
 Busqueda.Visible = True
 Busqueda.listDirs.Clear
 DestinoArch.Clear
 If unir = False Then DestinoDir.Clear
 Busqueda.listDirs.AddItem RutaInicial
 Y = 0
 Do Until Y = Busqueda.listDirs.ListCount
  DoEvents
    Busqueda.dirTemp.path = Busqueda.listDirs.List(Y)
    If Busqueda.dirTemp.ListCount > 0 Then
        For X = 0 To Busqueda.dirTemp.ListCount - 1
            Busqueda.listDirs.AddItem Busqueda.dirTemp.List(X)
            Busqueda.Label1 = Busqueda.dirTemp.List(X)
        Next X
    End If
    Y = Y + 1
 Loop
 For X = 0 To Busqueda.listDirs.ListCount - 1
  DoEvents
    If Busqueda.listDirs.List(X) Like "*\" Then
    txtfile = Dir(Busqueda.listDirs.List(X) & "*." & Trim(Tipo))
    Else
    txtfile = Dir(Busqueda.listDirs.List(X) & "\*." & Trim(Tipo))
    End If
    If Not txtfile = "" Then
        Do
            tfilename = path & txtfile
            If unir = True Then
             If Len(Busqueda.listDirs.List(X)) = 3 Then DestinoDir.AddItem Busqueda.listDirs.List(X) & tfilename
             If Len(Busqueda.listDirs.List(X)) > 3 Then DestinoDir.AddItem Busqueda.listDirs.List(X) & "\" & tfilename
            Else
             DestinoArch.AddItem tfilename
             If Len(Busqueda.listDirs.List(X)) = 3 Then DestinoDir.AddItem Busqueda.listDirs.List(X)
             If Len(Busqueda.listDirs.List(X)) > 3 Then DestinoDir.AddItem Busqueda.listDirs.List(X) & "\"
            End If
            txtfile = Dir$
        Loop Until txtfile = ""
    End If
 Next X
 Unload Busqueda
End Sub

Public Function FRACCIONAR(Texto As String, Valor As Integer) As String
 Dim n As Integer
 Dim X As Integer
 n = 1
 X = 0
 Do While n <= Len(Texto)
  If Mid(Texto, n, 1) <> "/" Then
   FRACCIONAR = FRACCIONAR + Mid(Texto, n, 1)
  Else
   X = X + 1
   If X = Valor Then
    Exit Do
   Else
    FRACCIONAR = ""
   End If
  End If
  n = n + 1
 Loop
End Function

Public Function UponSimbol(Texto As String, Simbolo As String, Direccion As Integer) As String
 Dim n As Integer
 Dim check As Boolean
 check = False
 n = 1
 UponSimbol = ""
 If Len(Trim(Simbolo)) > 1 Then Exit Function
 If Direccion = 0 Then
  Do While Mid(Texto, n, 1) <> Simbolo
   UponSimbol = UponSimbol + Mid(Texto, n, 1)
   n = n + 1
  Loop
 Else
  Do While n <= Len(Trim(Texto))
   If check = True Then UponSimbol = UponSimbol + Mid(Texto, n, 1)
   If Mid(Texto, n, 1) = Simbolo Then check = True
   n = n + 1
  Loop
 End If
End Function
