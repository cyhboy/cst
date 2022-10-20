
Private Function json_ParseArray(json_String As String, ByRef json_Index As Long) As Collection
    If testing Then Exit Function
    Set json_ParseArray = New Collection

    json_SkipSpaces json_String, json_Index
    If VBA.Mid$(json_String, json_Index, 1) <> "[" Then
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '['")
    Else
        json_Index = json_Index + 1

        Do
            json_SkipSpaces json_String, json_Index
            If VBA.Mid$(json_String, json_Index, 1) = "]" Then
                json_Index = json_Index + 1
                Exit Function
            ElseIf VBA.Mid$(json_String, json_Index, 1) = "," Then
                json_Index = json_Index + 1
                json_SkipSpaces json_String, json_Index
            End If

            json_ParseArray.Add json_ParseValue(json_String, json_Index)
        Loop
    End If
End Function

