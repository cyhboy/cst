
Public Function CntExeRunning(exeName As String) As Integer
    If testing Then Exit Function
    'On Error GoTo ErrorHandler
    On Error Resume Next
    'Dim flag As Boolean
    Dim cnt As Integer
    'cnt = 0
    Dim strComputer As String
    
    Dim objWMI As Object
    Dim objProcessSet As Object
    'Dim objProcess As Object
    
    Dim strUserName As String
    Dim strUserDomain As String
    
    strComputer = "."
    Set objWMI = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
    Set objProcessSet = objWMI.ExecQuery("SELECT Name FROM Win32_Process WHERE Name = '" & exeName & "'")
    'MsgBox objProcessSet.count
    
    cnt = objProcessSet.Count
    
    
'ErrorHandler:

    If Err.Number <> 0 Then
        'Do nothing as always error
        'MyMsgBox Err.Number & " " & Err.Description, 10
        'cnt = 0
    End If
    
    'MyMsgBox cnt & "", 10
    
    Set objProcessSet = Nothing
    Set objWMI = Nothing
    
    CntExeRunning = cnt
End Function

