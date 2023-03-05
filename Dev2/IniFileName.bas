
Public Function IniFileName() As String
    If testing Then
        Exit Function
    End If

    IniFileName = "C:\AppFiles\WinSCP\WinSCP.ini"
End Function

