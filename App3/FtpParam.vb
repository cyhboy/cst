
Public Sub FtpParam(hold As Boolean)
    If testing Then
        Exit Sub
    End If

    On Error GoTo ErrorHandler

    Call Fold

    Dim path As String
    Dim parameter As String
    Dim currentRow As Integer

    path = GetAppDrive() & "\FileZilla\filezilla.exe "
    currentRow = ActiveCell.Row

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

    parameter = "sftp://" & uid & ":" & pass & "@" & Cells(currentRow, 2) & ":" & port & fileOrFolder & " --local=""" & localFolder & """"
    'MsgBox path & parameter

    'Exit Sub
    ShellRunMax path & parameter

ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If

    If hold Then
        Dim exeName As String: exeName = ExtractEXE(path)
        While True = IsExeRunning(exeName)
            Sleep 10000
            ShellRun path & "--close", False
        Wend
    End If

End Sub

