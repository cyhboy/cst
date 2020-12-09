
Public Sub UnSltTitle()
    If testing Then Exit Sub
    Dim R As Range
    Dim RR As Range
    For Each R In Selection.Cells
        If R.Row <> 1 Then
            If RR Is Nothing Then
                Set RR = R
            Else
                Set RR = Application.Union(RR, R)
            End If
        End If
    Next R
    If Not RR Is Nothing Then
        RR.Select
    End If
End Sub

