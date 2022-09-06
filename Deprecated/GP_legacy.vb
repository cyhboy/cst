
Public Sub GP_legacy()
    If testing Then Exit Sub
    Dim videoPath As String
    Dim videoFileName As String
    Dim currentRow As Integer
    currentRow = ActiveCell.Row
    videoPath = Cells(currentRow, 9)
    videoFileName = Cells(currentRow, 11)
    
    Dim path As String
    path = "'" & GetAppDrive() & "\GP.ps1'"
    
    Dim parameter As String
    parameter = "'" & videoPath & "'" & " " & "'" & videoFileName & "'"
    Dim fullCommand As String
    fullCommand = path & " " & parameter
    ' MsgBox fullCommand
    PowerShellRun fullCommand, True
    
ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 20
    End If
End Sub

