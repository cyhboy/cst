
Public Sub Xx()
    If testing Then
        Exit Sub
    End If

    Dim currentRow As Integer
    currentRow = ActiveCell.Row

    Cells(currentRow, 24) = ""
End Sub

