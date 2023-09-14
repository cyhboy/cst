
Public Sub ColrR()
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
                RobotRunByParam "ColrR"
            End If
        Next curCell
        Exit Sub
    End If
    
    Dim lastCol, firstRow, lastRow As Long
    lastCol = ActiveSheet.[A1].End(xlToRight).Column
    firstRow = Selection.Cells(1, 1).Row
    lastRow = Selection.Cells(Selection.Rows.count, 1).Row
    Range(Cells(firstRow, 1), Cells(lastRow, lastCol)).Select
    With Selection.Interior
        'MsgBox .ColorIndex
        If IsNull(.ColorIndex) Or IsEmpty(.ColorIndex) Then
            .ColorIndex = 1
        End If
        If .ColorIndex >= 55 Or .ColorIndex < 1 Then
            .ColorIndex = 1
        End If

        .ColorIndex = .ColorIndex + 1

        While False = IsBackgroudColor(.Color) Or .ColorIndex < 3
            .ColorIndex = .ColorIndex + 1
        Wend
        'MsgBox .ColorIndex
    End With
End Sub

