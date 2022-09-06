
Public Sub PttyHOST()
    If testing Then
        Exit Sub
    End If

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

    Dim fqdn As String
    fqdn = Cells(currentRow, 2)

    Dim commandPath As String
    commandPath = GetBakDrive() & "\ptty_command.txt"

    Dim cmd As String
    cmd = "hostname -s" & Chr(13) & Chr(10) & "hostname -a" & Chr(13) & Chr(10) & "hostname -i" & Chr(13) & Chr(10) & "hostname -A" & Chr(13) & Chr(10) & "hostname -I"

    WriteTxt2Tmp cmd & Chr(13) & Chr(10) & "exit", commandPath

    parameter = fqdn & " -l " & uid & " -pw " & pass & " -m """ & commandPath & """ -t"

    If pass = "" Then
        WriteTxt2Tmp "dzdo /bin/su - " & uid & " -c '" & cmd & "'" & Chr(13) & Chr(10) & "exit", commandPath

        uid = Environ$("username")
        Dim propsMap As Variant
        Set propsMap = ReadPropertyInAppFiles("identity.ini")

        pass = propsMap("AD_PASSWORD")

        parameter = fqdn & " -l " & uid & " -pw " & pass & " -m """ & commandPath & """ -t"
    End If

    Dim cntEXE As Integer
    cntEXE = CntExeRunning(ExtractEXE(path))

    ShellRunHide path & parameter

    Sleep 1000

    Dim killFlag As Boolean
    While CntExeRunning(ExtractEXE(path)) - cntEXE > 0 And killFlag = False
        Sleep 3000
        If Now - LastModDate("C:\BAK\putty.log") > 3000 / 1000 / 60 / 24 Then
            'MyQuestionBox "Do U want to kill", "Yes", "No", 6
            'If confirmation = "Yes" Then
            'killFlag = KillExeRunning(ExtractEXE(path))
            'End If
        End If
    Wend

    'Sleep 1000

    '    Dim exeName As String: exeName = ExtractEXE(path)
    '    While True = IsExeRunning(exeName)
    '        Sleep 5000
    '    Wend

    Dim pttyResult1 As String
    pttyResult1 = SearchRegxKwInFile("C:\BAK\putty.log", "Using username .*\r\n(.*)\r\n", True)

    'Dim pttyResult1 As String
    'pttyResult1 = SearchRegxKwInFile("C:\BAK\putty.log", "(hk[^\.]*|tkcs[^\.]*|mtcs[^\.]*)$")

    'Dim pttyResult2 As String
    'pttyResult2 = searchRegxKwInFile("C:\BAK\putty.log", "(130[^ ]*)$")

    'Dim pttyResult3 As String
    'pttyResult3 = SearchRegxKwInFile("C:\BAK\putty.log", "(hk.*) $")

    'Dim pttyResult4 As String
    'pttyResult4 = SearchRegxKwInFile("C:\BAK\putty.log", "(130.*) $")

    'Dim arr1
    'Dim arr2
    'arr1 = Split(pttyResult3, " ")
    'arr2 = Split(pttyResult4, " ")

    '    If pttyResult <> "" Then
    '        pttyResult = "WebSphere MQ " & pttyResult
    '    End If

    Cells(currentRow, 1) = pttyResult1

    ' On Error Resume Next
    ' Cells(currentRow, 6) = arr2(Application.Match(fqdn, arr1, False) - 1)

    'If Err.Number <> 0 Then
    '    Cells(currentRow, 6) = arr2(0)
    'End If
ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If
End Sub

