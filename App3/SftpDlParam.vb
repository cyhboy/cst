
Public Sub SftpDlParam(hold As Boolean)
    If testing Then
        Exit Sub
    End If

    On Error GoTo ErrorHandler

    Dim path As String
    path = GetAppDrive() & "\FlashFXP\FlashFXP.exe "

    Dim parameter As String
    Dim currentRow As Integer
    currentRow = ActiveCell.Row

    Dim fqdn As String
    fqdn = Cells(currentRow, 2)

    Dim uid As String
    uid = Cells(currentRow, 3)

    Dim pass As String
    pass = Cells(currentRow, 4)

    Dim fileOrFolder As String
    fileOrFolder = Cells(currentRow, 5)

    Dim localFolder As String
    localFolder = Cells(currentRow, 9)

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

    Dim Length As Integer
    Length = Len(fileOrFolder)

    Dim index As Integer
    index = InStrRev(fileOrFolder, "/")


    fileOrFolder = Left(fileOrFolder, index)

    If pass = "" Then
        'uid = Environ$("username")

        Dim propsMap As Variant
        Set propsMap = ReadPropertyInAppFiles("identity.ini")

        pass = propsMap("AD_PASSWORD")
    End If

    If (InStr(Cells(currentRow, 11), "*") = 0 And InStr(Cells(currentRow, 11), "#") = 0) Then
        parameter = " -download sftp://" & uid & ":" & pass & "@" & fqdn & ":" & port & " -remotepath=""" & fileOrFolder & Cells(currentRow, 11) & """ -local=" & """" & localFolder & """" & " -localpath=" & """" & localFolder & Replace(Cells(currentRow, 11), "/", "\") & """"
    ElseIf True = StartsWith(Cells(currentRow, 11), "#") Then

        parameter = " -download sftp://" & uid & ":" & pass & "@" & fqdn & ":" & port & " -remotepath=""" & fileOrFolder & """ -local=" & """" & localFolder & """" & " -localpath=" & """" & localFolder & """" & " -ruleset=" & """" & Replace(Cells(currentRow, 11).Value, "#", "") & """"
    Else
        parameter = " -download sftp://" & uid & ":" & pass & "@" & fqdn & ":" & port & " -remotepath=""" & fileOrFolder & """ -local=" & """" & localFolder & """" & " -localpath=" & """" & localFolder & """" & " -ruleset=" & """" & Cells(currentRow, 11) & """"
    End If

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
            Sleep 5000
        Wend
    End If
End Sub

