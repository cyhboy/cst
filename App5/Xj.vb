
Public Sub Xj()
    If testing Then
        Exit Sub
    End If

    Dim currentRow As Integer
    currentRow = ActiveCell.Row

    Cells(currentRow, 10) = ""
End Sub

