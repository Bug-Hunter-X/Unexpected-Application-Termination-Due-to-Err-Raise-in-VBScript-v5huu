Function MyFunction(param1)
  On Error Resume Next
  If IsEmpty(param1) Then
    Err.Raise 13, , "Parameter cannot be empty"
  Else
    ' ... rest of the function
  End If
  If Err.Number <> 0 Then
    MsgBox "Error: " & Err.Description, vbCritical
    Err.Clear
  End If
End Function