
Public Sub InitCh()
    If testing Then
        Exit Sub
    End If

    Dim cell As Object
    For Each cell In Selection.Cells
        If cell.EntireColumn.Hidden = False And cell.EntireRow.Hidden = False Then
            cell.Value = "chrome.exe """""
        End If
    Next cell
End Sub

