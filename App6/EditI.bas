
Public Sub EditI()
    If testing Then
        Exit Sub
    End If

    Dim path As String
    path = "code.cmd"
    FileEditParam False, True, path
End Sub

