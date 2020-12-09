
Public Sub ShellRun(cmd As String)
    If testing Then Exit Sub
    On Error GoTo ErrorHandler
    Shell cmd, vbNormalFocus
    
    Exit Sub
ErrorHandler:
    If Not StartsWith(cmd, "cmd ") Then
        Err.Clear
        On Error Resume Next
        Shell "cmd /K " & cmd, vbNormalFocus
    End If
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If
End Sub

