
Public Sub ScpDlI()
    If testing Then
        Exit Sub
    End If

    ScpDlParam True
    Call XftpI
    If "On" = ReadRegAR() Then
        Dim currentRow As Integer
        currentRow = ActiveCell.Row
        Dim exer As String
        exer = Cells(currentRow, 16)
        If InStr(exer, "ScpDlI") = 0 Then
            Cells(currentRow, 16) = Trim(exer & " " & "ScpDlI")
        End If
    End If
End Sub

