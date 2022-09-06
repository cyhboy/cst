
Private Sub json_SkipSpaces(json_String As String, ByRef json_Index As Long)
    If testing Then Exit Sub
    ' Increment index to skip over spaces
    Do While json_Index > 0 And json_Index <= VBA.Len(json_String) And VBA.Mid$(json_String, json_Index, 1) = " "
        json_Index = json_Index + 1
    Loop
End Sub

