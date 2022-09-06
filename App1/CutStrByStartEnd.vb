
Public Function CutStrByStartEnd(orgStr As String, startStr As String, endStr As String, Optional includeStart As Boolean = False, Optional includeEnd As Boolean = False)
    If testing Then
        Exit Function
    End If

    Dim startPos As String
    Dim startLen As String

    If InStr(orgStr, startStr) > 0 Then
        startPos = InStr(orgStr, startStr) + Len(startStr)
        If includeStart Then
            startPos = InStr(orgStr, startStr)
        End If
        If InStr(startPos + 1, orgStr, endStr) > startPos Then
            startLen = InStr(startPos + 1, orgStr, endStr) - startPos

            If includeEnd Then
                startLen = startLen + Len(endStr)
            End If

            CutStrByStartEnd = Mid(orgStr, startPos, startLen)
        Else
            startLen = Len(orgStr) - startPos
            If startLen > 0 Then
                CutStrByStartEnd = Mid(orgStr, startPos, startLen + 1)
            End If
        End If
    Else
        CutStrByStartEnd = ""
    End If
End Function

