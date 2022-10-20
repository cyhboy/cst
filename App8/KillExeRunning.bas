
Public Function KillExeRunning(exeName As String, keepCnt As Integer) As Boolean
    If testing Then
        Exit Function
    End If

    On Error Resume Next
    Dim flag As Boolean
    flag = False

    Dim strComputer As String
    Dim objWMI As Object, objProcessSet As Object, objProcess As Object

    Dim strUserName As String
    Dim strUserDomain As String

    strComputer = "."
    Set objWMI = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
    Set objProcessSet = objWMI.ExecQuery("SELECT Name FROM Win32_Process WHERE Name = '" & exeName & "'")

    If objProcessSet.count > keepCnt Then

        For Each objProcess In objProcessSet

            objProcess.GetOwner strUserName, strUserDomain
            'MsgBox strUserName
            If strUserName = Environ$("username") Then
            End If
            'MsgBox "Process " & objProcess.Name & " is owned by " & strUserDomain & "\" & strUserName & "."

            If objProcess.Name = exeName Then
                Dim errReturnCode As Integer
                errReturnCode = objProcess.Terminate()
                'MsgBox errReturnCode
                If errReturnCode = 0 Then
                    flag = True
                    Exit For
                End If
            End If
        Next objProcess
    End If

    Set objProcessSet = Nothing
    Set objWMI = Nothing

    KillExeRunning = flag
End Function

