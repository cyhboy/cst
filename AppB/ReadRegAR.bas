
Public Function ReadRegAR() As String
    If testing Then Exit Function
    On Error GoTo ErrorHandler
    If recorder Then
        ReadRegAR = "On"
    Else
        ReadRegAR = "Off"
    End If
    
ErrorHandler:
    If Err.Number <> 0 Then
        'MyMsgBox Err.Number & " " & Err.Description, 30
        ReadRegAR = "Off"
    End If
End Function


