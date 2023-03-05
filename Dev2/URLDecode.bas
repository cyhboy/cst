
Public Function URLDecode(ByVal strEncodedURL As String) As String
    If testing Then
        Exit Function
    End If

    Dim str As String
    str = strEncodedURL
    If Len(str) > 0 Then
'        str = Replace(str, "&amp", " & ")
'        str = Replace(str, "&#03", Chr(39))
'        str = Replace(str, "&quo", Chr(34))
'        str = Replace(str, "+", " ")
'        str = Replace(str, "%2B", "+")
'        str = Replace(str, "%2A", "*")
'        str = Replace(str, "%40", "@")
'        str = Replace(str, "%2D", "-")
'        str = Replace(str, "%5F", "_")
'        str = Replace(str, "%2E", ".")
'        str = Replace(str, "%2F", "/")
        str = Replace(str, "%5C", "\")

        URLDecode = str
    End If

End Function

