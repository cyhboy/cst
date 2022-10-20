
Public Sub ShellRunStd(cmd As String)
    If testing Then
        Exit Sub
    End If

    Shell cmd, vbNormalFocus
End Sub

