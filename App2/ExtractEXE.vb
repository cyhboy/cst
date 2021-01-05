
Public Function ExtractEXE(fullPath As String)
    If testing Then Exit Function
    Dim tempLocation As Integer
    Dim tempStr As String
    
    tempLocation = InStr(fullPath, ".exe")
    
    If tempLocation = 0 Then
        'MsgBox "javaw.exe"
        ExtractEXE = "javaw.exe"
        Exit Function
    End If
    
    tempStr = Left(fullPath, tempLocation)

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

