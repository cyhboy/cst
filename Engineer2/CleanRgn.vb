
Public Sub CleanRgn()
    If testing Then
        Exit Sub
    End If
    ActiveCell.CurrentRegion.Select
    If Selection.Rows.count > 1 Then
        Call UnSltTitle
        Selection.ClearContents
    End If
End Sub

