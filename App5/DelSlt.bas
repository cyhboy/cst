
Public Sub DelSlt(Optional control As IRibbonControl)
    If testing Then
        Exit Sub
    End If

    Dim curCell As Range
    For Each curCell In Selection
        If curCell.EntireColumn.Hidden = False And curCell.EntireRow.Hidden = False Then
            curCell.Font.Strikethrough = Not curCell.Font.Strikethrough
        End If
    Next curCell
End Sub

