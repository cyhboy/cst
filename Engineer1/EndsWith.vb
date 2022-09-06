
Public Function EndsWith(str As String, ending As String) As Boolean
'    If testing Then
'        Exit Function
'    End If
    Dim endingLen As Integer
    endingLen = Len(ending)
    EndsWith = (Right(UCase(str), endingLen) = UCase(ending))
End Function

