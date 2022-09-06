
Private Function json_ParseKey(json_String As String, ByRef json_Index As Long) As String
    If testing Then Exit Function
    ' Parse key with single or double quotes
    If VBA.Mid$(json_String, json_Index, 1) = """" Or VBA.Mid$(json_String, json_Index, 1) = "'" Then
        json_ParseKey = json_ParseString(json_String, json_Index)
    ElseIf JsonOptions.AllowUnquotedKeys Then
        Dim json_Char As String
        Do While json_Index > 0 And json_Index <= Len(json_String)
            json_Char = VBA.Mid$(json_String, json_Index, 1)
            If (json_Char <> " ") And (json_Char <> ":") Then
                json_ParseKey = json_ParseKey & json_Char
                json_Index = json_Index + 1
            Else
                Exit Do
            End If
        Loop
    Else
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '""' or '''")
    End If

    ' Check for colon and skip if present or throw if not present
    json_SkipSpaces json_String, json_Index
    If VBA.Mid$(json_String, json_Index, 1) <> ":" Then
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting ':'")
    Else
        json_Index = json_Index + 1
    End If
End Function

