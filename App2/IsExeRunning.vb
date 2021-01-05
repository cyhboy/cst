
Public Function IsExeRunning(exeName As String) As Boolean
    If testing Then Exit Function
    On Error GoTo ErrorHandler
    
    Dim flag As Boolean
    Dim strComputer As String
    Dim objWMI As Object, objProcessSet As Object, objProcess As Object
    
    Dim strUserName As String
    Dim strUserDomain As String
    
    strComputer = "."
    Set objWMI = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
    Set objProcessSet = objWMI.ExecQuery("SELECT Name FROM Win32_Process WHERE Name = '" & exeName & "'")
    'MsgBox objProcessSet.count
    
    
'MsgBox Environ$("username")
For Each objProcess In objProcessSet
    objProcess.GetOwner strUserName, strUserDomain
    'MsgBox strUserName
    If strUserName = Environ$("username") Then
        flag = True
        Exit For
    End If
    'MsgBox "Process " & objProcess.Name & " is owned by " & strUserDomain & "\" & strUserName & "."
Next
    
    'If objProcessSet.count > 0 Then
    '    flag = True
    'Else
    '    flag = False
    'End If
    
'    For Each Process In objProcessSet
'        If Process.Name = exeName Then
'            flag = True
'            Exit For
'        End If
'    Next

ErrorHandler:
    Set objProcessSet = Nothing
    Set objWMI = Nothing
    
    If Err.Number <> 0 Then
        IsExeRunning = True
    Else
        IsExeRunning = flag
    End If
End Function

