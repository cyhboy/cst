
Public Function CutStringByStartAndEnd(orgStr As String, startStr As String, endStr As String)
    If testing Then
        Exit Function
    End If

    Dim startPos As String
    Dim startLen As String

    If InStr(orgStr, startStr) > 0 Then
        startPos = InStr(orgStr, startStr) + Len(startStr)
        If InStr(startPos, orgStr, endStr) > startPos Then
            startLen = InStr(startPos, orgStr, endStr) - startPos
            CutStringByStartAndEnd = Mid(orgStr, startPos, startLen)
        Else
            startLen = Len(orgStr) - startPos
            If startLen > 0 Then
                CutStringByStartAndEnd = Mid(orgStr, startPos, startLen + 1)
            End If
        End If
    Else
        CutStringByStartAndEnd = ""
    End If
End Function

