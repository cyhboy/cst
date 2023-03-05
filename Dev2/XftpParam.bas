
Public Sub XftpParam(Hold As Boolean)
    If testing Then
        Exit Sub
    End If

    On Error GoTo ErrorHandler

    Call Fold

    Dim path As String
    Dim parameter As String
    Dim currentRow As Integer

    path = GetAppDrive() & "\WinSCP\WinSCP.exe "
    currentRow = ActiveCell.Row

    Dim fqdn As String
    fqdn = Cells(currentRow, 2)

    Dim uid As String
    uid = Cells(currentRow, 3)
    Dim pass As String
    pass = Cells(currentRow, 4)

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

    Dim fileOrFolder As String
    fileOrFolder = Cells(currentRow, 5)

    Dim Length As Integer
    Length = Len(fileOrFolder)

    Dim index As Integer
    index = InStrRev(fileOrFolder, "/")

    fileOrFolder = Left(fileOrFolder, index)

    Dim localFolder As String
    localFolder = Cells(currentRow, 9)
    If EndsWith(localFolder, "\") Then
        localFolder = Left(localFolder, Len(localFolder) - 1)
    End If

    If pass = "" Then
        'uid = Environ$("username")

        Dim propsMap As Variant
        Set propsMap = ReadPropertyInAppFiles("identity.ini")

        pass = propsMap("AD_PASSWORD")
    End If

    parameter = "sftp://" & uid & ":" & pass & "@" & fqdn & ":" & port & fileOrFolder

    If ppkPath <> "" Then
        parameter = "sftp://" & uid & "@" & fqdn & ":" & port & fileOrFolder & " /privatekey=" & ppkPath
    End If

    If URLEncode(localFolder) <> ReadIniFileString("Configuration\Interface\Commander\LocalPanel", "LastPath") Then
        'MsgBox URLEncode(localFolder)
        'MsgBox ReadIniFileString("Configuration\Interface\Commander\LocalPanel", "LastPath")
        MsgBox WriteIniFileString("Configuration\Interface\Commander\LocalPanel", "LastPath", URLEncode(localFolder))
    End If

    'MsgBox path & parameter
    'Exit Sub

    ShellRunMax path & parameter

ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If

    '    If Hold Then
    '        Dim exeName As String: exeName = ExtractEXE(path)
    '        While True = IsExeRunning(exeName)
    '            Sleep 10000
    '            ShellRun path & "--close", False
    '        Wend
    '    End If

End Sub

