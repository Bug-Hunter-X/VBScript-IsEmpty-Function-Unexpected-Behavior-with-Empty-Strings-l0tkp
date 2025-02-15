Function MyFunction(param1)
  If param1 = vbNullString Then
    ' Handle empty string specifically
    param1 = ""
  ElseIf IsEmpty(param1) Then
    ' Handle uninitialized variables
    param1 = ""
  End If
  ' ... rest of the function
End Function