
Public Function ExtractEXE(FullPath As String)
    If testing Then
        Exit Function
    End If

    Dim tempLocation As Integer
    Dim tempStr As String

    tempLocation = InStr(FullPath, ".exe")

    If tempLocation = 0 Then
        'MsgBox "javaw.exe"
        ExtractEXE = "javaw.exe"
        Exit Function
    End If

    tempStr = Left(FullPath, tempLocation)

    If InStr(tempStr, """") > 0 Then
        tempStr = Right(tempStr, Len(tempStr) - InStrRev(tempStr, """"))
    End If

    If InStr(tempStr, " ") > 0 Then
        tempStr = Right(tempStr, Len(tempStr) - InStrRev(tempStr, " "))
    End If

    If InStr(tempStr, "\") = 0 Then
        tempStr = tempStr & "exe"
        'MsgBox tempStr
        ExtractEXE = tempStr
        Exit Function
    End If

    tempStr = Right(tempStr, Len(tempStr) - InStrRev(tempStr, "\"))

    tempStr = tempStr & "exe"
    'MsgBox tempStr
    ExtractEXE = tempStr
End Function

