
Public Sub RunAppRA()
    If testing Then
        Exit Sub
    End If
    Dim currentRow As Integer
    currentRow = ActiveCell.Row

    Cells(currentRow, 19) = "'" & Cells(currentRow, 18)
    Dim parameter As String
    parameter = Cells(currentRow, 10)

    Dim arr As Variant

    arr = Split(parameter, Chr(10))

    Dim path As String
    Dim i As Integer
    For i = 0 To UBound(arr)
        path = path & arr(i) & "&"
    Next i

    While Right(path, 1) = "&"
        path = Left(path, Len(path) - 1)
    Wend

    If Not Cells(currentRow, 9).HasFormula Then
        If Dir(Cells(currentRow, 9), vbDirectory) <> vbNullString Then
            Dim cdPath As String
            cdPath = Cells(currentRow, 9)
            path = "cd " & cdPath & "&" & path
        End If
    End If

    'MsgBox path

    Cells(currentRow, 18) = "'" & ShellRunResult(path, "C:\BAK\cmd.log", True)

    Cells(currentRow, 12) = LastModDate("C:\BAK\cmd.log")
End Sub


