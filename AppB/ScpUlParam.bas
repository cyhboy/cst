
Public Sub ScpUlParam(Hold As Boolean)
    If testing Then
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    Call Fold
    Dim path As String
    'path = "cmd.exe /C " & GetAppDrive() & "\WinSCP\WinSCP.com /script=WinSCP.txt"
    path = "cmd.exe /C " & GetAppDrive() & "\WinSCP\WinSCP.com /command"
    Dim parameter As String
    Dim currentRow As Integer
    currentRow = ActiveCell.Row
    Dim fqdn As String
    fqdn = Cells(currentRow, 2)
    Dim uid As String
    uid = Cells(currentRow, 3)
    Dim vla As String
    vla = Cells(currentRow, 3)
    Dim pass As String
    pass = Cells(currentRow, 4)
    Dim fileOrFolder As String
    fileOrFolder = Cells(currentRow, 5)
    Dim localFolder As String
    localFolder = Cells(currentRow, 9)
    Dim fileSet As String
    fileSet = Cells(currentRow, 11)
    Dim fileSetArr As Variant
    fileSetArr = Split(fileSet, Chr(10))
    Dim port As String
    port = Cells(currentRow, 7)
    If (port = "") Then
        port = "22"
    End If
    If (Len(port) > 5) Then
        port = "22"
    End If
    If Not IsNumeric(port) Then
        port = "22"
    End If
    
    Dim ppkPath As String: ppkPath = ""
    Dim ppkFile As String
    ppkFile = Cells(currentRow, 14)
    If EndsWith(ppkFile, ".ppk") Or ppkFile = "private_key" Then
        Dim ppkFolder As String
        ppkFolder = Cells(currentRow, 13)
        ppkPath = ppkFolder & ppkFile
    End If
    
    'Dim Length As Integer
    'Length = Len(fileOrFolder)
    Dim index As Integer
    index = InStrRev(fileOrFolder, "/")
    fileOrFolder = Left(fileOrFolder, index)
    If pass = "" Then
        uid = Environ$("username")
        Dim propsMap As Variant
        Set propsMap = ReadPropertyInAppFiles("identity.ini")
        pass = propsMap("AD_PASSWORD")
    End If
    'binary|ascii|automatic
    
    If ppkPath <> "" Then
        parameter = parameter & " " & """" & "open sftp://" & uid & "@" & fqdn & ":" & port & " -privatekey=" & ppkPath & """"
    Else
        parameter = parameter & " " & """" & "open sftp://" & uid & ":" & pass & "@" & fqdn & ":" & port & """"
    End If
    
    'parameter = parameter & " " & """" & "call dzdo /bin/su - " & vla & """"
    parameter = parameter & " " & """" & "cd " & fileOrFolder & """"
    parameter = parameter & " " & """" & "lcd """"" & localFolder & """"""""
    Dim i As Integer
    For i = 0 To UBound(fileSetArr)
        If InStr(fileSetArr(i), ".xls") > 0 Then
            parameter = parameter & " " & """" & "option transfer binary" & """"
        Else
            parameter = parameter & " " & """" & "option transfer ascii" & """"
        End If
        parameter = parameter & " " & """" & "put """"" & fileSetArr(i) & """"""""
    Next i
    parameter = parameter & " " & """" & "exit" & """"
    'MsgBox path & parameter
    'Exit Sub
    ShellRun path & parameter, True
ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If
    If Hold Then
        Dim exeName As String: exeName = ExtractEXE(path)
        While True = IsExeRunning(exeName)
            Sleep 5000
        Wend
    End If
End Sub

