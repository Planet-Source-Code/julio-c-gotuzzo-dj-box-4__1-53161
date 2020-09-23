Attribute VB_Name = "DJ4_Loader"

Function SaveDJ4(file As String) As Boolean
 Dim i As Integer
 SaveDJ4 = False
 On Error GoTo err
 Open file For Output As #1
   Print #1, "#DJBOX4 HEADER"
   Print #1, "[LISTASA]"
 If Form1.LISTA(1).ListCount > 0 Then
  For i = 0 To Form1.LISTA(1).ListCount - 1
    Print #1, Form1.LISTA(1).List(i)
    Print #1, Form1.LISTA(0).List(i)
    Print #1, Form1.ILISTA.List(i)
    Print #1, Form1.FLISTA.List(i)
  Next i
 End If
   Print #1, "[LISTASB]"
 If Form1.LISTB(1).ListCount > 0 Then
  For i = 0 To Form1.LISTB(1).ListCount - 1
    Print #1, Form1.LISTB(1).List(i)
    Print #1, Form1.LISTB(0).List(i)
    Print #1, Form1.ILISTB.List(i)
    Print #1, Form1.FLISTB.List(i)
  Next i
 End If
   Print #1, "[FX]"
 For i = 0 To Form4.EFFECT.ListCount - 1
   Print #1, Trim(Form4.FX(i).Caption)
   Print #1, Form4.EFFECT.List(i)
 Next i
   Print #1, "[RTM]"
 For i = 0 To Form4.RITMO.ListCount - 1
   Print #1, Trim(Form4.RTM(i).Caption)
   Print #1, Form4.RITMO.List(i)
 Next i
   Print #1, "[FIN]"
 Close #1
 SaveDJ4 = True
 Exit Function
err:
End Function

Function LoadDJ4(file As String) As Boolean
 Dim check As Integer
 Dim POSIT As Integer
 Dim n As Integer
 check = 0
 POSIT = 0
 LoadDJ4 = False
 Open file For Input As #1
 Do
  Line Input #1, a$
  If check = 0 And a$ = "#DJBOX4 HEADER" Then
   Form1.LISTA(1).Clear
   Form1.LISTA(0).Clear
   Form1.LISTB(1).Clear
   Form1.LISTB(0).Clear
   Form1.ILISTA.Clear
   Form1.ILISTB.Clear
   Form1.FLISTA.Clear
   Form1.FLISTB.Clear
   Form4.EFFECT.Clear
   Form4.RITMO.Clear
   check = 1
   POSIT = 0
  End If
  If check = 1 And a$ = "[LISTASA]" Then
   check = 2
   POSIT = 0
  End If
  If check = 2 And a$ = "[LISTASB]" Then
   check = 3
   POSIT = 0
  End If
  If check = 3 And a$ = "[FX]" Then
   check = 4
   POSIT = 0
   n = 0
  End If
  If check = 4 And a$ = "[RTM]" Then
   check = 5
   POSIT = 0
   n = 0
  End If
  If a$ = "[FIN]" Then Exit Do
  Select Case check
  Case Is = 2
   If a$ <> "[LISTASA]" Then
    Select Case POSIT
    Case Is = 0
     Form1.LISTA(1).AddItem a$
     POSIT = 1
    Case Is = 1
     Form1.LISTA(0).AddItem a$
     POSIT = 2
    Case Is = 2
     Form1.ILISTA.AddItem a$
     POSIT = 3
    Case Is = 3
     Form1.FLISTA.AddItem a$
     POSIT = 0
    End Select
   End If
  Case Is = 3
   If a$ <> "[LISTASB]" Then
    Select Case POSIT
    Case Is = 0
     Form1.LISTB(1).AddItem a$
     POSIT = 1
    Case Is = 1
     Form1.LISTB(0).AddItem a$
     POSIT = 2
    Case Is = 2
     Form1.ILISTB.AddItem a$
     POSIT = 3
    Case Is = 3
     Form1.FLISTB.AddItem a$
     POSIT = 0
    End Select
   End If
  Case Is = 4
   If a$ <> "[FX]" Then
    If POSIT = 0 Then
     Form4.FX(n).Caption = Trim(a$)
     POSIT = 1
     n = n + 1
    Else
     Form4.EFFECT.AddItem Trim(a$)
     POSIT = 0
    End If
   End If
  Case Is = 5
   If a$ <> "[RTM]" Then
    If POSIT = 0 Then
     Form4.RTM(n).Caption = Trim(a$)
     POSIT = 1
     n = n + 1
    Else
     Form4.RITMO.AddItem Trim(a$)
     POSIT = 0
    End If
   End If
  End Select
  If check = 0 Then Exit Function
 Loop While Not (EOF(1))
 Close #1
 LoadDJ4 = True
End Function

