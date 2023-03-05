
Public Function URLEncode(ByVal strDecodedURL As String) As String
    If testing Then
        Exit Function
    End If

    Dim str As String
    str = strDecodedURL
    If Len(str) > 0 Then
        str = Replace(str, "\", "%5C")
        
'        str = Replace(str, " & ", "&amp")
'        str = Replace(str, Chr(39), "&#03")
'        str = Replace(str, Chr(34), "&quo")
'        str = Replace(str, "+", "%2B")
'        str = Replace(str, " ", "+")
'        str = Replace(str, "*", "%2A")
'        str = Replace(str, "@", "%40")
'        str = Replace(str, "-", "%2D")
'        str = Replace(str, "_", "%5F")

'        str = Replace(str, ".", "%2E")
'        str = Replace(str, "/", "%2F")
        
        URLEncode = str
    End If

End Function

