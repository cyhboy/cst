
Public Sub Robn()
    If testing Then Exit Sub
    
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
                RobotRunByParam "Robn"
            End If
        Next curCell
        Exit Sub
    End If
    
    Dim currentRow As Integer
    currentRow = ActiveCell.Row
    
    Dim actions As String
    actions = Cells(currentRow, 16)
    
    Dim action As String
    
    If CntSubstring(actions, " #") = 1 Then
        action = CutStringByStartAndEnd(actions, " #", " #")
        If InStr(action, " ") = 0 Then
            actions = Replace(actions, " FXplr", " #FXplr")
            Call Robot
            RobotRunByParam action
        End If
    Else
        
    End If
End Sub

