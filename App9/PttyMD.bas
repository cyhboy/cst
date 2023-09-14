
Public Sub PttyMD()
    If testing Then
        Exit Sub
    End If

    On Error GoTo ErrorHandler

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
                RobotRunByParam "PttyMD"
            End If
        Next curCell
        Exit Sub
    End If

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

End Sub

