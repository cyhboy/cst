
Public Sub Sample()
    If testing Then
        Exit Sub
    End If

    '    Dim n As Integer
    '    n = Selection.Count
    '    If n > 1 Then
    '        n = Selection.SpecialCells(xlCellTypeVisible).Count
    '    End If
    '    If n > 1 Then
    '        Dim curCell As Range
    '        For Each curCell In Selection
    '            If curCell.EntireColumn.Hidden = False And curCell.EntireRow.Hidden = False Then
    '                curCell.Select
    '                'MsgBox subName
    '                RobotRunByParam "Sample"
    '            End If
    '        Next curCell
    '        Exit Sub
    '    End If

    Dim currentRow As Integer
    currentRow = ActiveCell.Row

    If Trim(Cells(currentRow, 14)) = "" Then
        Cells(currentRow, 14).FormulaR1C1 = "=RIGHT(CELL(""filename"", R1C1),LEN(CELL(""filename"", R1C1))-FIND(""]"",CELL(""filename"", R1C1)))"
    End If

    If Trim(Cells(currentRow, 12)) = "" Then
        Cells(currentRow, 12).FormulaR1C1 = "=TODAY()"
    End If

    If Trim(Cells(currentRow, 15)) = "" Then
        Cells(currentRow, 15) = "'" & LPad((currentRow - 1) & "", 4, "0")
    End If

    If Trim(Cells(currentRow, 9)) = "" Then
        Cells(currentRow, 9).FormulaR1C1 = "=""C:\Deploy\"" & RC[-1] &  ""\"" & RC[-7] & ""_"" & RC[-8] & ""\"""
    End If

    If Trim(Cells(currentRow, 1)) = "" Then
        Cells(currentRow, 1) = "U"
    End If

    If Trim(Cells(currentRow, 2)) = "" Then
        Cells(currentRow, 2) = "U"
    End If

    If Trim(Cells(currentRow, 8)) = "" Then
        Cells(currentRow, 8) = "Task Memo"
    End If

    If Trim(Cells(currentRow, 16)) = "" Then
        Cells(currentRow, 16) = "MyTest"
    End If
End Sub

