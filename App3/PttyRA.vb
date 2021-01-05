
Public Sub PttyRA()
    If testing Then Exit Sub
    
    'On Error GoTo ErrorHandler
    
    Dim n As Integer
    n = Selection.Count
    If n > 1 Then
        n = Selection.SpecialCells(xlCellTypeVisible).Count
    End If
    If n > 1 Then
        Dim curCell As Range
        For Each curCell In Selection
            If curCell.EntireColumn.Hidden = False And curCell.EntireRow.Hidden = False Then
                curCell.Select
                'MsgBox subName
                RobotRunByParam "PttyRA"
            End If
        Next curCell
        Exit Sub
    End If
    
    Dim path As String
    path = GetAppDrive() & "\ptty\putty.exe "
    
    Dim parameter As String
    
    Dim currentRow As Integer
    currentRow = ActiveCell.Row
   
    Dim uid As String
    uid = Cells(currentRow, 3)
    
    Dim pass As String
    pass = Cells(currentRow, 4)
    
    Dim fqdn As String
    fqdn = Cells(currentRow, 2)
    
    Dim port As String
    port = Cells(currentRow, 7)
    
    Cells(currentRow, 19) = "'" & Cells(currentRow, 18)
    
    If port = "" Then
        'port = "2200"
        port = "22"
    End If
    
    If pass = "" Then
        Dim propsMap As Variant
        Set propsMap = ReadPropertyInAppFiles("identity.ini")

        pass = propsMap("AD_PASSWORD")
    End If
    
    Dim commandPath As String
    commandPath = GetBakDrive() & "\ptty_command.txt"
    
    Dim cmd As String

    cmd = Cells(currentRow, 10)
    
    WriteTxt2Tmp cmd & Chr(13) & Chr(10) & "exit", commandPath
    
    parameter = fqdn & " -l " & uid & " -pw " & pass & " -P " & port & " -m """ & commandPath & """ -t"
    
    Dim cntEXE As Integer
    cntEXE = CntExeRunning(ExtractEXE(path))
    
    ShellRunStd path & parameter
    
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
    
    Sleep 1000
    
    Dim pttyResult As String

    pttyResult = ReadLineByFile("C:\BAK\putty.log")
    
    Cells(currentRow, 18) = "'" & pttyResult
    
    Cells(currentRow, 12) = LastModDate("C:\BAK\putty.log")
    
'ErrorHandler:
'    If Err.Number <> 0 Then
'        MyMsgBox Err.Number & " " & Err.Description, 10
'    End If
End Sub

