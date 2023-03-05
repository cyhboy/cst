
Public Sub PowerShellRun(cmd As String, Hold As Boolean)
    If testing Then
        Exit Sub
    End If
    On Error Resume Next
    ' Shell "powershell.exe -ExecutionPolicy Unrestricted -File " & cmd, vbNormal

    ' MsgBox "powershell.exe -ExecutionPolicy Unrestricted -noexit " & """" & "& " & cmd
    ' Shell "cmd.exe /c powershell.exe -ExecutionPolicy Unrestricted -noexit " & """" & "& " & cmd, vbNormal

    Shell "powershell.exe -ExecutionPolicy Unrestricted -noexit " & """" & "& " & cmd, vbNormal

    ' KillExeRunning ExtractEXE("powershell.exe"), 2
    ' KillExeRunning ExtractEXE("conhost.exe"), 1

    '    Dim exeName As String: exeName = ExtractEXE("cmd.exe")
    '    While True = IsExeRunning(exeName)
    '        Sleep 3000
    '    Wend

    If Err.Number <> 0 Then
        MsgBox "Met a unexpected case: " & Err.Number
    End If
End Sub

