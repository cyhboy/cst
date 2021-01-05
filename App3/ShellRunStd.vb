
Public Sub ShellRunStd(cmd As String)
    If testing Then Exit Sub
    
    Shell cmd, vbNormalFocus
End Sub

