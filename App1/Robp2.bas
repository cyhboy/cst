
Public Sub Robp2()
    If testing Then
        Exit Sub
    End If

    Dim n As Integer
    n = Selection.count
    If n > 1 Then
        n = Selection.SpecialCells(xlCellTypeVisible).count
        'offsetCnt = 0
    End If

    If n > 1 Then
        Dim curCell As Range
        'Dim curRng As Range
        'Set curRng = Selection
        'Dim i As Integer
        'For i = 1 To curRng.count

        'MsgBox CurCell.Address
        For Each curCell In Selection
            'Set curCell = curRng(i, 0)
            If curCell.EntireColumn.Hidden = False And curCell.EntireRow.Hidden = False Then
                'MsgBox subName
                curCell.Select
                'If EndsWith(Cells(curCell.row, 16), " Dr") And i > 1 Then
                'offsetCnt = offsetCnt - 1
                'Else

                'End If
                RobotRunByParam "Robp"
            End If
        Next curCell
        'Next

        Exit Sub
    End If

    'If Not EndsWith(Cells(ActiveCell.row, 16), " Dr") And Cells(ActiveCell.row, 16) <> "" Then
    '    offsetCnt = 0
    'End If

    'ActiveCell.Offset(offsetCnt, 0).Select

    Dim currentRow As Integer
    currentRow = ActiveCell.Row

    Dim actions As String
    actions = Cells(currentRow, 16)

    Dim action As String

    If StartsWith(actions, "#") And InStr(actions, " ") > 0 Then
        action = CutStringByStartAndEnd(actions, "#", " ")
        RobotRunByParam action
        Call Robot
    Else

    End If

End Sub

