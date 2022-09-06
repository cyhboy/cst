
Public Sub RobotRunByParam(comm As String)
    If testing Then
        Exit Sub
    End If
    ' On Error GoTo ErrorHandler
    ' Application.Run comm
    Dim comms As Variant
    comms = Split(comm, "_")
    Dim i As Integer
    For i = 0 To UBound(comms)
        Application.Run comms(i)
    Next i
    ' ErrorHandler:
    '    If Err.Number <> 0 Then
    '        MyMsgBox Err.Number & " " & Err.Description & " " & comm, 30
    '    End If
End Sub

