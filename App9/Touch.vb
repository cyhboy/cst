
Public Sub Touch()
    If testing Then
        Exit Sub
    End If
    'On Error GoTo ErrorHandler
    Dim currentRow As Integer
    currentRow = ActiveCell.Row
    
    Dim fileName As String
    fileName = Cells(currentRow, 11)
    Dim localFolder As String
    localFolder = Cells(currentRow, 9)
    
    Dim filePath As String
    filePath = localFolder & fileName
    If InStr(fileName, ".doc") > 0 Then
        Call TouchDoc
        Exit Sub
    End If
    
    Dim wb As Workbook

    Dim appExcel As New Application
    appExcel.Visible = False
    appExcel.EnableEvents = False
    
    Set wb = appExcel.Workbooks.Open(filePath)
    
    wb.Save
    wb.Close savechanges:=True
    appExcel.Quit
    Set appExcel = Nothing
    
ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 10
    End If
End Sub

