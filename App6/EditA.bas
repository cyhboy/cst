
Public Sub EditA()
    If testing Then
        Exit Sub
    End If

    Dim path As String
    path = """" & GetAppDrive() & "\EditPlus\editplus.exe" & """ -e"
    ' path = """" & GetAppDrive() & "\EditPlus\editplus.exe"
    FileEditParam False, False, path
End Sub

