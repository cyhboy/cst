
Public Sub DelFsCll()
    If testing Then
        Exit Sub
    End If

    MyQuestionBox "delete file in cell? ", "No", "Yes", 10
    If confirmation = "No" Then
        Exit Sub
    End If

    On Error GoTo ErrorHandler
    Dim path As String
    Dim parameter As String
    path = "C:\AppFiles\cmdutils\Recycle.exe -f "
    'path = "Recycle.exe "
    Dim cell As Object
    For Each cell In Selection.Cells
        If cell.EntireColumn.Hidden = False And cell.EntireRow.Hidden = False Then
            parameter = " " & """" & Replace(cell.Value, Chr(10), """ """) & """"
            ShellRun path & parameter, False
        End If
    Next cell

ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If
End Sub



