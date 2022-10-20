
Public Function InUse(filePath As String) As Boolean
    If testing Then
        Exit Function
    End If
    Dim ff As Integer
    On Error Resume Next
    Open filePath For Binary Access Read Lock Read As #ff
    Close #ff
    InUse = IIf(Err.Number > 0, True, False)
    On Error GoTo 0
End Function

