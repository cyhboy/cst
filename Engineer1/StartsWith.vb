
Public Function StartsWith(str As String, start As String) As Boolean
    'If testing Then Exit Function
    str = CStr(str)
    start = CStr(start)
    Dim startLen As Integer
    startLen = Len(start)
    StartsWith = (Left(UCase(str), startLen) = UCase(start))
End Function

