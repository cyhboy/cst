
Public Sub DelWb()
    If testing Then
        Exit Sub
    End If

    MyQuestionBox "delete activated workbook in row? ", "No", "Yes", 10
    If confirmation = "No" Then
        Exit Sub
    End If

    Dim path As String
    Dim parameter As String
    'path = "cmd.exe /C C:\AppFiles\cmdutils\Recycle -f "
    path = "C:\AppFiles\cmdutils\Recycle.exe -f "
    'path = "Recycle.exe "

    parameter = ActiveWorkbook.FullName
    
    ActiveWorkbook.Close
    ShellRun path & parameter, False

    Dim exeName As String: exeName = ExtractEXE(path)
    While True = IsExeRunning(exeName)
        Sleep 3000
    Wend
End Sub

