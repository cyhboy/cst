
Private Function json_BufferToString(ByRef json_Buffer As String, ByVal json_BufferPosition As Long) As String
    If testing Then Exit Function
    If json_BufferPosition > 0 Then
        json_BufferToString = VBA.Left$(json_Buffer, json_BufferPosition)
    End If
End Function

