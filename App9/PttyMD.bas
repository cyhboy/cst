
Public Sub PttyMD()
    If testing Then
        Exit Sub
    End If

    '    On Error GoTo LineHandler
    '    Dim subName As String
    '362:     Err.Raise 1979
    'LineHandler:
    '    If Err.Number = 1979 Then
    '        subName = GetSubName("AllSpecial", Erl)
    '        If subName = "" Then
    '            MsgBox "Process name is not available. Please contact administrator. "
    '            'Exit Sub
    '        Else
    '            LogToDB subName
    '        End If
    '    End If
    '    Resume Next

    On Error GoTo ErrorHandler
    Dim path As String

    Dim parameter As String
    Dim currentRow As Integer

    path = GetAppDrive() & "\ptty\putty.exe "
    currentRow = ActiveCell.Row


    Dim uid As String
    uid = Cells(currentRow, 3)

    Dim pass As String
    pass = Cells(currentRow, 4)

    Dim commandPath As String
    commandPath = GetBakDrive() & "\ptty_command.txt"

    Dim cmd As String
    cmd = "mkdir -p " & Cells(currentRow, 5)

    WriteTxt2Tmp cmd & Chr(13) & Chr(10) & "exit", commandPath

    parameter = Cells(currentRow, 2) & " -l " & uid & " -pw " & pass & " -m """ & commandPath & """ -t"

    If pass = "" Then
        WriteTxt2Tmp "dzdo /bin/su - " & uid & " -c '" & cmd & "'" & Chr(13) & Chr(10) & "exit", commandPath

        uid = Environ$("username")
        Dim propsMap As Variant
        Set propsMap = ReadPropertyInAppFiles("identity.ini")
        pass = propsMap("AD_PASSWORD")

        parameter = Cells(currentRow, 2) & " -l " & uid & " -pw " & pass & " -m """ & commandPath & """ -t"
    End If

    ShellRunHide path & parameter

    Dim exeName As String: exeName = ExtractEXE(path)
    While True = IsExeRunning(exeName)
        Sleep 5000
    Wend

    Dim pttyResult As String
    pttyResult = SearchRegxKwInFile("C:\BAK\putty.log", "(fail)")

    If pttyResult = "" Then
        MsgBox "Server Folder Created Successfully"
    Else
        MsgBox "Server Folder Fail To Create"
    End If

ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If

    '    If "On" = ReadRegAR() Then
    '        Dim exer As String
    '        exer = Cells(currentRow, 16)
    '        If InStr(exer, subName) = 0 Then
    '            Cells(currentRow, 16) = Trim(exer & " " & subName)
    '        End If
    '    End If
End Sub

