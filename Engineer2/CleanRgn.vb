
Public Sub CleanRgn()
    If testing Then Exit Sub
    ActiveCell.CurrentRegion.Select
    If Selection.Rows.Count > 1 Then
        Call UnSltTitle
        Selection.ClearContents
    End If
End Sub

