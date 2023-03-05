
Public Sub PttyRunParam(Hold As Boolean)
    If testing Then
        Exit Sub
    End If

    On Error GoTo ErrorHandler
    Dim path As String
    Dim currentRow As Integer
    path = GetAppDrive() & "\ptty\putty.exe "
    currentRow = ActiveCell.Row
    Dim parameter As String

    Dim fqdn As String
    fqdn = Cells(currentRow, 2)

    Dim uid As String
    uid = Cells(currentRow, 3)

    If uid = "" Then
        uid = Environ$("username")
    End If

    Dim pass As String
    pass = Cells(currentRow, 4)

    Dim remoteFolder As String
    remoteFolder = Cells(currentRow, 5)

    Dim port As String
    port = Cells(currentRow, 7)

    If port = "" Then
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

    If Trim(pass) = "" Then
        'ruid = Environ$("username")
        Dim propsMap As Variant
        Set propsMap = ReadPropertyInAppFiles("identity.ini")
        pass = propsMap("AD_PASSWORD")
        'If Trim(uid) <> "" Then
        '    WriteTxt2Tmp "dzdo /bin/su - " & uid & " -c '" & cmd & "'", commandPath
        'End If
    End If

    Dim commandPath As String
    commandPath = GetBakDrive() & "\ptty_command.txt"

    Dim cmd As String
    cmd = "pwd" & Chr(13) & Chr(10) & "set -x" & Chr(13) & Chr(10) & Cells(currentRow, 10)

    If Trim(remoteFolder) <> "" Then
        'WriteTxt2Tmp "cd " & remoteFolder & Chr(13) & Chr(10) & cmd & Chr(13) & Chr(10) & "exit", commandPath
        WriteTxt2Tmp "cd " & remoteFolder & Chr(13) & Chr(10) & "pwd" & Chr(13) & Chr(10) & cmd & Chr(13) & Chr(10) & "/bin/bash", commandPath
    Else
        WriteTxt2Tmp cmd & Chr(13) & Chr(10) & "/bin/bash", commandPath
    End If

    parameter = fqdn & " -l " & uid & " -pw " & pass & " -P " & port & " -m """ & commandPath & """ -t"

    If ppkPath <> "" Then
        parameter = fqdn & " -l " & uid & " -i """ & ppkPath & """ -P " & port & " -m """ & commandPath & """ -t"
    End If

    Dim cntEXE As Integer
    cntEXE = CntExeRunning(ExtractEXE(path))

    'MsgBox path & parameter
    'Exit Sub

    ShellRunMax path & parameter

    If Hold Then
        While CntExeRunning(ExtractEXE(path)) - cntEXE > 0
            Sleep 3000
        Wend
    End If

ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If

End Sub

