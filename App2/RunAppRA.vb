
Public Sub RunAppRA()
    If testing Then Exit Sub
    Dim currentRow As Integer
    currentRow = ActiveCell.Row
    
    Cells(currentRow, 19) = "'" & Cells(currentRow, 18)
    Dim parameter As String
    parameter = Cells(currentRow, 10)
    
    Dim arr
    
    arr = Split(parameter, Chr(10))
    
    Dim path As String
    Dim i
    For i = 0 To UBound(arr)
        path = path & arr(i) & "&"
    Next
    
    While Right(path, 1) = "&"
        path = Left(path, Len(path) - 1)
    Wend
    
    If Not Cells(currentRow, 9).HasFormula Then
        Dim cdPath As String
        cdPath = Cells(currentRow, 9)
        path = "cd " & cdPath & "&" & path
    End If
    

    Cells(currentRow, 18) = "'" & ShellRunResult(path, "C:\BAK\cmd.log", True)
    
    Cells(currentRow, 12) = LastModDate("C:\BAK\cmd.log")
End Sub

