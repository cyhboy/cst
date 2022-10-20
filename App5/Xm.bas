
Public Sub Xm()
    If testing Then
        Exit Sub
    End If

    Dim currentRow As Integer
    currentRow = ActiveCell.Row
    Cells(currentRow, 13) = ""
End Sub

