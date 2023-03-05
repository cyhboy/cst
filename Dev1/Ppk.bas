
Public Sub Ppk()
    If testing Then
        Exit Sub
    End If
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
                RobotRunByParam "Ppk"
            End If
        Next curCell
        Exit Sub
    End If
    
    Dim currentRow As Integer
    currentRow = ActiveCell.Row
    
    Dim ppkPath As String: ppkPath = ""
    Dim ppkFile As String
    ppkFile = Cells(currentRow, 14)
    If EndsWith(ppkFile, ".ppk") Or ppkFile = "private_key" Then
        Dim ppkFolder As String
        ppkFolder = Cells(currentRow, 13)
        ppkPath = ppkFolder & ppkFile
    End If
    
    Dim inputKey As String
    Dim outputKey As String
    
    If EndsWith(ppkPath, ".ppk") Then
        inputKey = Replace(ppkPath, ".ppk", "")
        outputKey = ppkPath
    Else
        inputKey = ppkPath
        outputKey = ppkPath & ".ppk"
    End If

    Dim cmdStr As String

    cmdStr = "C:\AppFiles\WinSCP\winscp.com /keygen " & inputKey & " /output=" & outputKey

    ShellRun cmdStr, True
End Sub

