
Public Sub ShellRun(cmd As String, isKeep As Boolean)
    If testing Then
        Exit Sub
    End If
    On Error GoTo ErrorHandler

    If isKeep Then
        If Not StartsWith(cmd, "cmd ") Then
            Shell "cmd /K " & cmd, vbNormalFocus
        End If
    Else
        Shell cmd, vbNormalFocus
    End If

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

