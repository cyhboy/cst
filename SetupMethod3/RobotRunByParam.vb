
Public Sub RobotRunByParam(comm As String)
    If testing Then Exit Sub
    'On Error GoTo ErrorHandler
    Application.Run comm
'ErrorHandler:
'    If Err.Number <> 0 Then
'        MyMsgBox Err.Number & " " & Err.Description & " " & comm, 30
'    End If
End Sub

