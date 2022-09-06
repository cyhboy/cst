
Public Sub ChOp()
    If testing Then
        Exit Sub
    End If

    Dim path As String
    Dim parameter As String
    Dim currentRow As Integer
    currentRow = ActiveCell.Row
    'path = "firefox.exe "
    path = "chrome.exe "
    parameter = Cells(currentRow, 10)
    If InStr(parameter, "http") > 0 Then
        parameter = CutStrByStartEnd(parameter, "http", "$", True)
    Else
        parameter = ""
    End If
    'MsgBox path & """" & parameter & """"
    ShellRunStd path & """" & parameter & """"
End Sub

