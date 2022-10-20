
Public Sub TagProcRun(resultStr As String, keyword As String, isRegx As Boolean, isMulti As Boolean, colNumber As Integer)
    If testing Then
        Exit Sub
    End If
    
    Dim currentRow As Integer
    currentRow = ActiveCell.Row
    
    If isRegx Then
        If MatchRegx(resultStr, keyword) Then
            If isMulti Then
                Cells(currentRow, colNumber) = Join(SearchRegxKwInStrMultToList(resultStr, keyword, 0, True), Chr(10))
            Else
                Cells(currentRow, colNumber) = SearchRegxKwInStr(resultStr, keyword, True)
            End If
        End If
    End If
End Sub

