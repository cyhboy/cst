
Public Sub WriteTxt2Tmp(text As String, path As String)
    If testing Then
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    Dim ff As Integer
    ff = FreeFile()
    Open path For Output As #ff
    Print #ff, text
    Close #ff
ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If
End Sub

