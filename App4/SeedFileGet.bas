
Public Function SeedFileGet(filePath As String) As String
    If testing Then
        Exit Function
    End If

    Dim seedFilePath As String
    Dim filename As String
    Dim ext As String
    If InStr(filePath, ".") > 0 Then
        ext = Right(filePath, Len(filePath) - InStrRev(filePath, ".") + 1)
    End If
    If InStr(filePath, "_") > 0 Then
        filename = Left(filePath, InStrRev(filePath, "_") - 1)
    End If
    SeedFileGet = filename & ext
End Function


