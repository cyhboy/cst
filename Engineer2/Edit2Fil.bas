
Public Sub Edit2Fil()
    If testing Then
        Exit Sub
    End If
    Dim currentRow As Integer
    currentRow = ActiveCell.Row
    Dim module As String
    module = Cells(currentRow, 1)
    Dim subb As String
    subb = Cells(currentRow, 2)

    Dim path As String
    Dim parameter As String
    path = """" & GetAppDrive() & "\EditPlus\editplus.exe"" "
    parameter = "C:\SANDBOX\VB_SPACE\VBA_PROJECT\" & Format(Now, "yyyyMMdd") & "\" & module & "\" & subb & ".bas"
    ShellRun path & parameter, False
End Sub

