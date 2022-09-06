
Public Sub Rd2J()
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
                RobotRunByParam "Rd2J"
            End If
        Next curCell
        Exit Sub
    End If

    Dim currentRow As Integer
    currentRow = ActiveCell.Row

    Dim lineCnt As Integer
    lineCnt = CntSubstring(Cells(currentRow, 10), Chr(10))

    Cells(currentRow, 23) = "'" & lineCnt

ErrorHandler:
    If Err.Number <> 0 Then
        PrintResult "FAILED"
        MyMsgBox Err.Number & " " & Err.Description, 10
    Else
        If lineCnt > 1 Then
            PrintResult "NOT"
        Else
            PrintResult "END"
        End If
    End If
End Sub

