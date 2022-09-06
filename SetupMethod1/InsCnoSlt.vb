
Public Sub InsCnoSlt()
    If testing Then
        Exit Sub
    End If

    Dim cell As Object

    For Each cell In Selection.Cells
        If InStr(cell.Value, ") ") = 2 Or InStr(cell.Value, ") ") = 3 Then
            cell.Value = Right(cell.Value, Len(cell.Value) - InStr(cell.Value, ") ") - 2 + 1)
        End If

        cell.Value = cell.Column & ") " & cell.Value

    Next cell

End Sub

