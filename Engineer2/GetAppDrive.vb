
Public Function GetAppDrive() As String
    If testing Then Exit Function
    Dim appDrive As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    appDrive = fso.GetDriveName(ActiveWorkbook.path) & "\AppFiles"
    
    If InStr(appDrive, ":") = 0 Then
        appDrive = "C:" & "\AppFiles"
    End If
    
    If InStr(appDrive, "D:") > 0 Or InStr(appDrive, "d:") > 0 Then
        appDrive = "C:" & "\AppFiles"
    End If
    
    Set fso = Nothing
    
    GetAppDrive = appDrive
End Function

