
Public Sub Xy()
    If testing Then
        Exit Sub
    End If

    Dim currentRow As Integer
    currentRow = ActiveCell.Row

    Cells(currentRow, 25) = ""
End Sub

