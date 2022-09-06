
Public Function LPad(str As String, strLen As Integer, padStr As String) As String
    If testing Then
        Exit Function
    End If
    Dim n: n = 0
    If strLen > Len(str) Then
        n = strLen - Len(str)
    End If
    'LPad = String(n, padStr) & str
    LPad = Replace(Space(n), " ", padStr) & str
End Function

