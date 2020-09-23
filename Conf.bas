Attribute VB_Name = "Config_Module"
Option Explicit

Function INIGETVALUE(Texto As String, Field As String) As String
    Dim searchtext, tBuf
    
    searchtext = InStr(Texto, Field & "=")
    searchtext = searchtext - 1
    tBuf = right(Texto, Len(Texto) - searchtext)
    searchtext = InStr(tBuf, Chr(13))
    searchtext = searchtext - 1
    tBuf = left(tBuf, searchtext)
    tBuf = right(tBuf, Len(tBuf) - Len(Field) - 1)
    INIGETVALUE = tBuf
End Function

Function INISETVALUE(Texto As String, Field As String, Value As String)
    'Change value
     Texto = Replace(Texto, Field & "=" & INIGETVALUE(Texto, Field), Field & "=" & Value)
End Function

Function LOADINI(Texto As String, Filename As String) As Boolean
    'Load an ini file into the text variable
    Open Filename For Input As #1
        Texto = Input$(LOF(1), #1)
    Close #1
    If Texto <> "" Then LOADINI = True Else LOADINI = False
End Function

Function SAVEINI(Texto As String, Filename As String)
    'Save an ini file from text variable
    Open Filename For Output As #1
        Print #1, Texto
    Close #1
End Function

Function INIADDVALUE(Texto As String, Variable As String, Value As String)
    Texto = Texto & Variable & "=" & Value & vbCrLf
End Function

