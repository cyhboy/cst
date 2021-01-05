
Public Function GetBakDrive() As String
    If testing Then Exit Function
    Dim bakDrive As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    bakDrive = fso.GetDriveName(ActiveWorkbook.path) & "\BAK"
    If InStr(bakDrive, ":") = 0 Then
        bakDrive = "C:" & "\BAK"
    End If
    
    If InStr(bakDrive, "D:") > 0 Or InStr(bakDrive, "d:") > 0 Then
        bakDrive = "C:" & "\BAK"
    End If
    Set fso = Nothing
    GetBakDrive = bakDrive
End Function

