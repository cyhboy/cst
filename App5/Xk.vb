
Public Sub Xk()
    If testing Then
        Exit Sub
    End If

    Dim currentRow As Integer
    currentRow = ActiveCell.Row

    Cells(currentRow, 11) = ""
End Sub

