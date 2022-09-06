
Public Sub Xt()
    If testing Then
        Exit Sub
    End If

    Dim currentRow As Integer
    currentRow = ActiveCell.Row
    Cells(currentRow, 20) = ""
End Sub

