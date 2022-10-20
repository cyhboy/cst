
Public Sub PingIp()
    If testing Then
        Exit Sub
    End If

    Dim path As String
    Dim parameter As String
    Dim currentRow As Integer

    path = "cmd.exe /K Ping -a "
    currentRow = ActiveCell.Row

    'Ping -a 130.29.48.148
    parameter = Cells(currentRow, 6)

    ShellRun path & parameter, False
End Sub

