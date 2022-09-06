
Public Sub PttyRegx()
    If testing Then
        Exit Sub
    End If
    'On Error GoTo ErrorHandler

    Dim n As Integer
    n = Selection.count
    If n > 1 Then
        n = Selection.SpecialCells(xlCellTypeVisible).count
    End If
    If n > 1 Then
        Dim curCell As Range
        For Each curCell In Selection
            If curCell.EntireColumn.Hidden = False And curCell.EntireRow.Hidden = False Then
                curCell.Select
                'MsgBox subName
                RobotRunByParam "PttyRegx"
            End If
        Next curCell
        Exit Sub
    End If

    Dim path As String
    Dim parameter As String
    Dim currentRow As Integer

    path = GetAppDrive() & "\ptty\putty.exe "
    currentRow = ActiveCell.Row


    Dim fqdn As String
    fqdn = Cells(currentRow, 2)

    Dim uid As String
    uid = Cells(currentRow, 3)

    Dim pass As String
    pass = Cells(currentRow, 4)

    Dim port As String
    port = Cells(currentRow, 7)

    Dim command As String
    command = Cells(currentRow, 10)
    command = Replace(command, Chr(10), Chr(13) & Chr(10))

    Dim commandPath As String
    commandPath = GetBakDrive() & "\ptty_command.txt"

    WriteTxt2Tmp command & Chr(13) & Chr(10) & "exit", commandPath
    'Exit Sub
    parameter = fqdn & " -l " & uid & " -pw " & pass & " -P " & port & " -m """ & commandPath & """ -t"

    If Trim(pass) = "" Then

        'WriteTxt2Tmp "dzdo /bin/su - " & uid & " -c '" & command & "'" & Chr(13) & Chr(10) & "exit", commandPath

        'uid = Environ$("username")

        Dim propsMap As Variant
        Set propsMap = ReadPropertyInAppFiles("identity.ini")

        pass = propsMap("AD_PASSWORD")

        parameter = fqdn & " -l " & uid & " -pw " & pass & " -P " & port & " -m """ & commandPath & """ -t"
    End If

    Dim cntEXE As Integer
    cntEXE = CntExeRunning(ExtractEXE(path))

    ShellRunHide path & parameter

ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 10
    End If

    While CntExeRunning(ExtractEXE(path)) - cntEXE > 0
        Sleep 3000
    Wend

    Dim pttyResult As String
    'pttyResult = SearchRegxKwInFileMultToListLast("C:\BAK\putty.log", Cells(currentRow, 20))

    pttyResult = SearchRegxKwInFile("C:\BAK\putty.log", Cells(currentRow, 20), True)

    Cells(currentRow, 17) = pttyResult

End Sub

