
Public Sub DelFs()
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
                RobotRunByParam "DelFs"
            End If
        Next curCell
        Exit Sub
    End If

    MyQuestionBox "delete file in row? ", "No", "Yes", 10
    If confirmation = "No" Then
        Exit Sub
    End If

    Dim path As String
    Dim parameter As String
    'path = "cmd.exe /C C:\AppFiles\cmdutils\Recycle -f "
    path = "C:\AppFiles\cmdutils\Recycle.exe -f "
    'path = "Recycle.exe "

    Dim currentRow As Integer

    currentRow = ActiveCell.Row

    parameter = """" & Cells(currentRow, 9) & Cells(currentRow, 11) & """"

    ShellRun path & parameter, False

    Dim exeName As String: exeName = ExtractEXE(path)
    While True = IsExeRunning(exeName)
        Sleep 3000
    Wend
End Sub

