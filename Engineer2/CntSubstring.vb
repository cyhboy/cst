
Public Function CntSubstring(text As String, subStr As String, Optional ignoreFlag As Boolean = False) As Long
    If testing Then Exit Function
    
    If ignoreFlag Then
        CntSubstring = (Len(UCase(text)) - Len(Replace(UCase(text), UCase(subStr), ""))) / Len(UCase(subStr))
    Else
        CntSubstring = (Len(text) - Len(Replace(text, subStr, ""))) / Len(subStr)
    End If
End Function

