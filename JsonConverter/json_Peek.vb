
Private Function json_Peek(json_String As String, ByVal json_Index As Long, Optional json_NumberOfCharacters As Long = 1) As String
    If testing Then Exit Function
    ' "Peek" at the next number of characters without incrementing json_Index (ByVal instead of ByRef)
    json_SkipSpaces json_String, json_Index
    json_Peek = VBA.Mid$(json_String, json_Index, json_NumberOfCharacters)
End Function

