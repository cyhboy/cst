
Public Sub PingResult()
    If testing Then
        Exit Sub
    End If

    Dim path As String
    Dim parameter As String
    Dim currentRow As Integer

    path = "cmd.exe /K Ping -a "
    currentRow = ActiveCell.Row

    parameter = Cells(currentRow, 2)

    Dim PingResult As String
    PingResult = GetPingResult(parameter)

    Dim objWMI As Object
    Dim response1 As Object
    Dim r1 As Object

    If PingResult = "Connected" Then
        Set objWMI = GetObject("winmgmts:")
        Set response1 = objWMI.ExecQuery(" Select * from Win32_PingStatus WHERE address='" & parameter & "'")
        For Each r1 In response1
            'MsgBox "DNS Name:" & r1.Address & " has addresses: " & r1.ProtocolAddress
            If Cells(currentRow, 6).Value = "" Or Cells(currentRow, 6).Value = r1.ProtocolAddress Then
                Cells(currentRow, 6).Value = r1.ProtocolAddress
            Else
                Cells(currentRow, 6).Value = Cells(currentRow, 6).Value & Chr(10) & r1.ProtocolAddress
            End If
        Next r1
        Cells(currentRow, 5).Value = "ONLINE"
    Else
        If PingResult = "Request timed out" Then
            Call LookupIp
        Else
            Cells(currentRow, 6) = ""
        End If

        Cells(currentRow, 5).Value = "DEMISED"
    End If
End Sub

