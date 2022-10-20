
Public Sub FXplr()
    If testing Then
        Exit Sub
    End If
    Dim path As String
    Dim parameter As String
    path = "explorer "

    Dim currentRow As Integer

    Dim cell As Object
    For Each cell In Selection.Cells
        If cell.EntireColumn.Hidden = False And cell.EntireRow.Hidden = False Then

            currentRow = cell.Row

            If InStr(Cells(currentRow, 11), ".") > 0 Then
                parameter = " " & """" & Cells(currentRow, 9) & Cells(currentRow, 11) & """"
            Else
                parameter = " " & """" & Cells(currentRow, 9) & Replace(Right(Cells(currentRow, 10), Len(Cells(currentRow, 10)) - InStrRev(Cells(currentRow, 10), "/")), """", "") & """"
            End If
            'MsgBox parameter
            ShellRun path & parameter, False
            'ShellRunWait path & parameter
        End If
    Next cell

    Dim sleeping As Variant
    sleeping = Cells(currentRow, 22)

    If IsNumeric(sleeping) Then
        If CInt(sleeping) > 20000 Then
            sleeping = 20000
        End If
    Else
        sleeping = 0
    End If
    Sleep CInt(sleeping)
End Sub

