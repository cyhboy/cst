
Public Sub ShellRunMax(cmd As String)
    If testing Then
        Exit Sub
    End If
    'vbHide 0 Window is hidden and focus is passed to the hidden window. The vbHide constant is not applicable on Macintosh platforms.
    'vbNormalFocus 1 Window has focus and is restored to its original size and position.
    'vbMinimizedFocus 2 Window is displayed as an icon with focus.
    'vbMaximizedFocus 3 Window is maximized with focus.
    'vbNormalNoFocus 4 Window is restored to its most recent size and position. The currently active window remains active.
    'vbMinimizedNoFocus 6 Window is displayed as an icon. The currently active window remains active.
    On Error GoTo ErrorHandler
    Shell cmd, vbMaximizedFocus
ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If
End Sub

