
Public Sub ShellRunHide(cmd As String)
    If testing Then
        Exit Sub
    End If
    'On Error GoTo ErrorHandler
    Shell cmd, vbHide
    'ErrorHandler:
    '    If Err.Number <> 0 Then
    '        MyMsgBox Err.Number & " " & Err.Description, 30
    '    End If
End Sub

