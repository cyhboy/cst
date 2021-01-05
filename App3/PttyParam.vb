
Public Sub PttyParam(Hold As Boolean)
    If testing Then Exit Sub
    On Error GoTo ErrorHandler
    
    Dim path As String
    Dim parameter As String
    Dim currentRow As Integer
    path = GetAppDrive() & "\ptty\putty.exe "
    currentRow = ActiveCell.Row
    
    Dim fqdn As String
    fqdn = Cells(currentRow, 2)
    
    Dim uid As String
    uid = Cells(currentRow, 3)
    If uid = "" Then
        uid = Environ$("username")
    End If
    
    Dim pass As String
    pass = Cells(currentRow, 4)
    
    Dim port As String
    port = Cells(currentRow, 7)
    
    If port = "" Then
        port = "22"
    End If
    
    If Trim(pass) = "" Then
        'ruid = Environ$("username")
        'ruid = uid
        
        Dim propsMap As Variant
        Set propsMap = ReadPropertyInAppFiles("identity.ini")
        pass = propsMap("AD_PASSWORD")
        'MsgBox pass
        'If Trim(uid) <> "" Then
            'WriteTxt2Tmp "dzdo /bin/su - " & uid, commandPath
            'parameter = fqdn & " -l " & uid & " -pw " & pass & " -P " & port & " -m """ & commandPath & """ -t"
        'Else
            'parameter = fqdn & " -l " & uid & " -pw " & pass & " -P " & port
        'End If
    End If
    
    'Dim commandPath As String
    'commandPath = GetBakDrive() & "\ptty_command.txt"
    
    'WriteTxt2Tmp cmd & Chr(13) & Chr(10) & "exit", commandPath
    'parameter = fqdn & " -l " & uid & " -pw " & pass & " -P " & port & " -m """ & commandPath & """ -t"
    
    parameter = fqdn & " -l " & uid & " -pw " & pass & " -P " & port
    
    Dim cntEXE As Integer
    cntEXE = CntExeRunning(ExtractEXE(path))
    
    'MsgBox path & parameter
    'Exit Sub
    
    'ShellRunMax path & parameter
    ShellRun path & parameter
    
ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 10
    End If
    
    If Hold Then
        While CntExeRunning(ExtractEXE(path)) - cntEXE > 0
            Sleep 3000
        Wend
    End If
End Sub

