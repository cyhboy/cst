
Public Function LPad(str As String, strLen As Integer, padStr As String) As String
    If testing Then Exit Function
    Dim n: n = 0
    If strLen > Len(str) Then n = strLen - Len(str)
    LPad = String(n, padStr) & str
End Function

