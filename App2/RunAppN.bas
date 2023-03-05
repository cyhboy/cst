
Public Sub RunAppN()
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
                RobotRunByParam "RunAppN"
            End If
        Next curCell
        Exit Sub
    End If
    
    RunAppParam False, False, True, , "next"
End Sub

