Function f(x)
  If IsNumeric(x) Then
    If x < 0 Then
      f = -1
    ElseIf x = 0 Then
      f = 0
    Else
      f = 1
    End If
  Else
    f = -2 ' Handle non-numeric input
  End If
End Function

MsgBox f(-1) 
MsgBox f("abc")
MsgBox f(true) 